package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.model.AbbreviatedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.MorphologicalChart;
import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.openxml.builder.wml.WmlAdapter;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.nio.file.Path;
import java.util.concurrent.Callable;

import static org.apache.commons.lang3.ArrayUtils.isEmpty;

/**
 * @author sali
 */
public class MorphologicalChartAdapter extends DocumentAdapter implements Callable<Boolean> {

    private final Path path;
    private final ChartConfiguration chartConfiguration;
    private final MorphologicalChart[] charts;

    /**
     * Creates new document based on default configuration and charts.
     *
     * @param charts {@link MorphologicalChart} given charts.
     */
    public MorphologicalChartAdapter(MorphologicalChart... charts) {
        this(null, charts);
    }

    /**
     * Creates new document based on configuration and charts.
     *
     * @param chartConfiguration {@link ChartConfiguration} to how document to be rendered.
     * @param charts             {@link MorphologicalChart} given charts.
     */
    public MorphologicalChartAdapter(ChartConfiguration chartConfiguration, MorphologicalChart... charts) {
        this(null, chartConfiguration, charts);
    }

    /**
     * Creates new document based on configuration and charts.
     *
     * @param path               Name of file to save.
     * @param chartConfiguration {@link ChartConfiguration} to how document to be rendered.
     * @param charts             {@link MorphologicalChart} given charts.
     */
    public MorphologicalChartAdapter(Path path, ChartConfiguration chartConfiguration, MorphologicalChart... charts) {
        this.path = path;
        this.chartConfiguration = (chartConfiguration == null) ? new ChartConfiguration() : chartConfiguration;
        this.charts = charts;
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
        if (isEmpty(charts)) {
            return;
        }
        if (!chartConfiguration.isOmitAbbreviatedConjugation() && !chartConfiguration.isOmitToc()) {
            try {
                WmlAdapter.addTableOfContent(mdp, "Table Of Contents", " TOC \\o \"1-3\" \\h \\z \\t \"Arabic-Heading1,1\" ");
            } catch (Docx4JException e) {
                e.printStackTrace();
            }
        }
        for (MorphologicalChart mc : charts) {
            buildMorphologicalChart(mdp, mc);
        }
    }

    private void buildMorphologicalChart(MainDocumentPart mdp, MorphologicalChart mc) {
        final AbbreviatedConjugation abbreviatedConjugation = mc.getAbbreviatedConjugation();
        final boolean omitAbbreviatedConjugation = (abbreviatedConjugation == null) || chartConfiguration.isOmitAbbreviatedConjugation();
        if (!omitAbbreviatedConjugation) {
            AbbreviatedConjugationAdapter aca = new AbbreviatedConjugationAdapter(chartConfiguration, mc.getAbbreviatedConjugation());
            aca.buildDocument(mdp);
        }
        DetailedConjugationAdapter dca = new DetailedConjugationAdapter(mc.getDetailedConjugation());
        dca.buildDocument(mdp);
    }
}
