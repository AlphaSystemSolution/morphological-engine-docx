package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationData;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationTemplate;
import com.alphasystem.morphologicalengine.model.AbbreviatedConjugation;
import com.alphasystem.morphologicalengine.model.DetailedConjugation;
import com.alphasystem.morphologicalengine.model.MorphologicalChart;
import com.alphasystem.openxml.builder.wml.TocGenerator;
import com.alphasystem.openxml.builder.wml.WmlAdapter;
import com.alphasystem.openxml.builder.wml.WmlBuilderFactory;
import org.apache.commons.io.FilenameUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.P;

import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * @author sali
 */
public class MorphologicalChartEngine extends DocumentAdapter {

    private final AbbreviatedConjugationFactory abbreviatedConjugationFactory;
    private final DetailedConjugationFactory detailedConjugationFactory;
    private final SupplierFactory supplierFactory;
    private final ConjugationTemplate conjugationTemplate;

    MorphologicalChartEngine(AbbreviatedConjugationFactory abbreviatedConjugationFactory,
                             DetailedConjugationFactory detailedConjugationFactory,
                             SupplierFactory supplierFactory,
                             ConjugationTemplate conjugationTemplate) {
        this.abbreviatedConjugationFactory = abbreviatedConjugationFactory;
        this.detailedConjugationFactory = detailedConjugationFactory;
        this.supplierFactory = supplierFactory;
        this.conjugationTemplate = conjugationTemplate;
    }

    public void createDocument(Path path) throws Docx4JException {
        final ChartConfiguration chartConfiguration = (conjugationTemplate == null) ? new ChartConfiguration() :
                conjugationTemplate.getChartConfiguration();
        WmlHelper.createDocument(path, chartConfiguration, this);
        if (!chartConfiguration.isOmitAbbreviatedConjugation() && !chartConfiguration.isOmitDetailedConjugation()) {
            ChartConfiguration ch = new ChartConfiguration(chartConfiguration)
                    .omitDetailedConjugation(true)
                    .omitToc(true);
            WmlHelper.createDocument(toPath(path, "Abbreviated"), ch,
                    new MorphologicalChartEngine(abbreviatedConjugationFactory, detailedConjugationFactory, supplierFactory,
                            new ConjugationTemplate(conjugationTemplate).withChartConfiguration(ch)));

            ch = ch.omitAbbreviatedConjugation(true).omitDetailedConjugation(false);
            WmlHelper.createDocument(toPath(path, "Detailed"), ch,
                    new MorphologicalChartEngine(abbreviatedConjugationFactory, detailedConjugationFactory, supplierFactory,
                    new ConjugationTemplate(conjugationTemplate).withChartConfiguration(ch)));
        }
    }

    @Override
    protected void buildDocument(MainDocumentPart mdp) {
        if (conjugationTemplate == null) {
            return;
        }
        final ChartConfiguration chartConfiguration = conjugationTemplate.getChartConfiguration();
        final boolean addDetailedConjugation = !chartConfiguration.isOmitDetailedConjugation();
        final boolean addToc = !chartConfiguration.isOmitAbbreviatedConjugation() && !chartConfiguration.isOmitToc();
        final String tocHeading = "Table of Contents";
        final String bookmarkName = tocHeading.replaceAll(" ", "_").toLowerCase();
        if (addToc) {
            new TocGenerator()
                    .tocHeading(tocHeading)
                    .mainDocumentPart(mdp)
                    .instruction(" TOC \\o \"1-3\" \\h \\z \\t \"Arabic-Heading1,1\" ")
                    .tocStyle("TOCArabic")
                    .generateToc();
        }
        final List<MorphologicalChart> charts = createMorphologicalCharts();
        if (!charts.isEmpty()) {
            addToDocument(mdp, chartConfiguration, charts.get(0));
            charts.stream().skip(1).forEach(morphologicalChart -> {
                if (addToc) {
                    addBackLink(mdp, bookmarkName);
                }
                if (addDetailedConjugation) {
                    mdp.addObject(WmlAdapter.getPageBreak());
                }
                addToDocument(mdp, chartConfiguration, morphologicalChart);
            });
        }
    }

    private void addToDocument(MainDocumentPart mdp, ChartConfiguration chartConfiguration, MorphologicalChart morphologicalChart) {
        final AbbreviatedConjugation abbreviatedConjugation = morphologicalChart.getAbbreviatedConjugation();
        final boolean omitAbbreviatedConjugation = (abbreviatedConjugation == null) || chartConfiguration.isOmitAbbreviatedConjugation();
        if (!omitAbbreviatedConjugation) {
            AbbreviatedConjugationAdapter aca = abbreviatedConjugationFactory.createAbbreviatedConjugationAdapter(
                    chartConfiguration, morphologicalChart.getAbbreviatedConjugation());
            aca.buildDocument(mdp);
        }

        final DetailedConjugation detailedConjugation = morphologicalChart.getDetailedConjugation();
        final boolean omitDetailedConjugation = (detailedConjugation == null) || chartConfiguration.isOmitDetailedConjugation();
        if (!omitDetailedConjugation) {
            DetailedConjugationAdapter dca = detailedConjugationFactory.createDetailedConjugationAdapter(detailedConjugation);
            dca.buildDocument(mdp);
        }
    }

    private void addBackLink(MainDocumentPart mdp, String bookmarkName) {
        final P.Hyperlink backLink = WmlAdapter.addHyperlink(bookmarkName, "Back to Top");
        final P p = WmlBuilderFactory.getPBuilder().addContent(backLink).getObject();
        mdp.addObject(p);
    }

    public List<MorphologicalChart> createMorphologicalCharts() {
        final List<ConjugationData> data = conjugationTemplate.getData();
        final List<MorphologicalChart> morphologicalCharts = new ArrayList<>(data.size());
        for (ConjugationData conjugationData : data) {
            MorphologicalChartSupplier supplier = supplierFactory.createSupplier(conjugationData);
            morphologicalCharts.add(supplier.get());
        }
        return morphologicalCharts;
    }

    /**
     * Convert given path to new path by appending given suffix in the name of file.
     * <p>
     * This method is used to convert given path to create abbreviated and detail conjugation files.
     *
     * @param src    source path
     * @param suffix suffix to attach
     * @return new path
     */
    private Path toPath(final Path src, final String suffix) {
        final Path parent = src.getParent();
        final String fileName = src.getFileName().toString();
        final String baseName = FilenameUtils.getBaseName(fileName);
        final String extension = FilenameUtils.getExtension(fileName);
        return Paths.get(parent.toAbsolutePath().toString(), String.format("%s-%s.%s", baseName, suffix, extension));
    }

}
