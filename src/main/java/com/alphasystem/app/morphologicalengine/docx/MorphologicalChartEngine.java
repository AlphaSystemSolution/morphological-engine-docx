package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.model.AbbreviatedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.DetailedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.MorphologicalChart;
import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationData;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationTemplate;
import com.alphasystem.openxml.builder.wml.WmlAdapter;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;

/**
 * @author sali
 */
public class MorphologicalChartEngine extends DocumentAdapter implements Callable<Boolean> {

    private final Path path;
    private final ConjugationTemplate conjugationTemplate;

    /**
     * Creates new document based on default configuration and charts.
     *
     * @param conjugationTemplate {@link ConjugationTemplate} given charts.
     */
    public MorphologicalChartEngine(ConjugationTemplate conjugationTemplate) {
        this(null, conjugationTemplate);
    }

    /**
     * Creates new document based on template.
     *
     * @param path                Name of file to save.
     * @param conjugationTemplate {@link ConjugationTemplate} given template.
     */
    public MorphologicalChartEngine(Path path, ConjugationTemplate conjugationTemplate) {
        this.path = path;
        this.conjugationTemplate = conjugationTemplate;
    }

    @Override
    public Boolean call() throws Exception {
        if (path == null) {
            throw new IllegalArgumentException("Path cannot be null.");
        }
        createDocument(path);
        return true;
    }

    @Override
    protected void buildDocument(MainDocumentPart mdp) {
        if (conjugationTemplate == null) {
            return;
        }
        final ChartConfiguration chartConfiguration = conjugationTemplate.getChartConfiguration();
        if (!chartConfiguration.isOmitAbbreviatedConjugation() && !chartConfiguration.isOmitToc()) {
            try {
                WmlAdapter.addTableOfContent(mdp, "Table Of Contents", " TOC \\o \"1-3\" \\h \\z \\t \"Arabic-Heading1,1\" ");
            } catch (Docx4JException e) {
                e.printStackTrace();
            }
        }
        final List<MorphologicalChart> charts = createMorphologicalCharts();
        charts.forEach(morphologicalChart -> addToDocument(mdp, chartConfiguration, morphologicalChart));
    }

    private void addToDocument(MainDocumentPart mdp, ChartConfiguration chartConfiguration, MorphologicalChart morphologicalChart) {
        final AbbreviatedConjugation abbreviatedConjugation = morphologicalChart.getAbbreviatedConjugation();
        final boolean omitAbbreviatedConjugation = (abbreviatedConjugation == null) || chartConfiguration.isOmitAbbreviatedConjugation();
        if (!omitAbbreviatedConjugation) {
            AbbreviatedConjugationAdapter aca = new AbbreviatedConjugationAdapter(chartConfiguration, morphologicalChart.getAbbreviatedConjugation());
            aca.buildDocument(mdp);
        }

        final DetailedConjugation detailedConjugation = morphologicalChart.getDetailedConjugation();
        final boolean omitDetailedConjugation = (detailedConjugation == null) || chartConfiguration.isOmitDetailedConjugation();
        if (!omitDetailedConjugation) {
            DetailedConjugationAdapter dca = new DetailedConjugationAdapter(detailedConjugation);
            dca.buildDocument(mdp);
        }
    }

    public List<MorphologicalChart> createMorphologicalCharts() {
        final ChartConfiguration chartConfiguration = conjugationTemplate.getChartConfiguration();
        final List<ConjugationData> data = conjugationTemplate.getData();
        final List<MorphologicalChart> morphologicalCharts = new ArrayList<>(data.size());
        for (ConjugationData conjugationData : data) {
            MorphologicalChartSupplier supplier = new MorphologicalChartSupplier(chartConfiguration, conjugationData);
            morphologicalCharts.add(supplier.get());
        }
        return morphologicalCharts;
    }

}
