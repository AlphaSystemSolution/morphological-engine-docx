package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationData;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationTemplate;
import com.alphasystem.morphologicalengine.model.AbbreviatedConjugation;
import com.alphasystem.morphologicalengine.model.DetailedConjugation;
import com.alphasystem.morphologicalengine.model.MorphologicalChart;
import com.alphasystem.openxml.builder.wml.TocGenerator;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

/**
 * @author sali
 */
public class MorphologicalChartEngine extends DocumentAdapter {

    private ConjugationTemplate conjugationTemplate;

    public void setConjugationTemplate(ConjugationTemplate conjugationTemplate) {
        this.conjugationTemplate = conjugationTemplate;
    }

    public MorphologicalChartEngine conjugationTemplate(ConjugationTemplate conjugationTemplate) {
        setConjugationTemplate(conjugationTemplate);
        return this;
    }

    public void createDocument(Path path) throws Docx4JException {
        final ChartConfiguration chartConfiguration = (conjugationTemplate == null) ? new ChartConfiguration() :
                conjugationTemplate.getChartConfiguration();
        WmlHelper.createDocument(path, chartConfiguration, this);
    }

    @Override
    protected void buildDocument(MainDocumentPart mdp) {
        if (conjugationTemplate == null) {
            return;
        }
        final ChartConfiguration chartConfiguration = conjugationTemplate.getChartConfiguration();
        if (!chartConfiguration.isOmitAbbreviatedConjugation() && !chartConfiguration.isOmitToc()) {
            TocGenerator tocGenerator = new TocGenerator().tocHeading("Table of Contents").mainDocumentPart(mdp)
                    .instruction(" TOC \\o \"1-3\" \\h \\z \\t \"Arabic-Heading1,1\" ").tocStyle("TOCArabic");
            tocGenerator.generateToc();
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
        final List<ConjugationData> data = conjugationTemplate.getData();
        final List<MorphologicalChart> morphologicalCharts = new ArrayList<>(data.size());
        for (ConjugationData conjugationData : data) {
            MorphologicalChartSupplier supplier = new MorphologicalChartSupplier(conjugationData);
            morphologicalCharts.add(supplier.get());
        }
        return morphologicalCharts;
    }

}
