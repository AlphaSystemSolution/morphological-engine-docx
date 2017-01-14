package com.alphasystem.app.morphologicalengine.docx.test;

import com.alphasystem.app.morphologicalengine.conjugation.model.AbbreviatedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.DetailedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.MorphologicalChart;
import com.alphasystem.app.morphologicalengine.docx.AbbreviatedConjugationAdapter;
import com.alphasystem.app.morphologicalengine.docx.DetailedConjugationAdapter;
import com.alphasystem.app.morphologicalengine.docx.MorphologicalChartEngine;
import com.alphasystem.app.morphologicalengine.docx.WmlHelper;
import com.alphasystem.arabic.model.NamedTemplate;
import com.alphasystem.arabic.ui.util.FontUtilities;
import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationData;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationTemplate;
import com.alphasystem.morphologicalanalysis.morphology.model.RootLetters;
import com.alphasystem.morphologicalanalysis.morphology.model.support.VerbalNoun;
import org.apache.commons.lang3.ArrayUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import static com.alphasystem.arabic.model.ArabicLetterType.*;
import static com.alphasystem.arabic.model.NamedTemplate.*;
import static com.alphasystem.morphologicalanalysis.morphology.model.support.VerbalNoun.VERBAL_NOUN_V1;
import static java.lang.String.format;
import static java.lang.System.getProperty;
import static java.nio.file.Paths.get;
import static org.testng.Assert.*;
import static org.testng.Reporter.log;

/**
 * @author sali
 */
public class MorphologicalChartEngineTest {

    private static Path parentDocDir = null;

    static {
        parentDocDir = get(getProperty("target.dir"), "docs");
        try {
            Files.createDirectories(parentDocDir);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void openFile(Path path) {
        try {
            Desktop.getDesktop().open(path.toFile());
        } catch (IOException e) {
            // ignore
        }
    }

    @Test
    public void testCreateEmptyDocument() {
        final Path path = get(parentDocDir.toString(), "mydoc.docx");
        log(format("File Path: %s", path), true);
        MorphologicalChartEngine morphologicalChartEngine = new MorphologicalChartEngine(null, null);
        try {
            morphologicalChartEngine.createDocument(path);
            Assert.assertEquals(Files.exists(path), true);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
    }

    @Test(dependsOnMethods = {"testCreateEmptyDocument"})
    public void runConjugationBuilder() {
        final Path path = get(parentDocDir.toString(), "conjugations.docx");
        MorphologicalChartEngine morphologicalChartEngine = new MorphologicalChartEngine(getConjugationTemplate(getChartConfiguration()));
        try {
            morphologicalChartEngine.createDocument(path);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        openFile(path);
    }

    @Test(dependsOnMethods = {"runConjugationBuilder"})
    public void buildAbbreviatedConjugations() {
        ChartConfiguration chartConfiguration = getChartConfiguration().omitToc(true).omitDetailedConjugation(true);
        final ConjugationTemplate conjugationTemplate = getConjugationTemplate(chartConfiguration);
        MorphologicalChartEngine engine = new MorphologicalChartEngine(conjugationTemplate);
        List<MorphologicalChart> charts = engine.createMorphologicalCharts();
        final MorphologicalChart chart = charts.get(0);
        assertNotNull(chart);
        assertNull(chart.getDetailedConjugation());

        final Path path = get(parentDocDir.toString(), "abbreviated-conjugations.docx");

        AbbreviatedConjugation[] abbreviatedConjugations = new AbbreviatedConjugation[charts.size()];
        for (int i = 0; i < charts.size(); i++) {
            abbreviatedConjugations[i] = charts.get(i).getAbbreviatedConjugation();
        }
        AbbreviatedConjugationAdapter aca = new AbbreviatedConjugationAdapter(chartConfiguration, abbreviatedConjugations);
        try {
            WmlHelper.createDocument(path, chartConfiguration, aca);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        openFile(path);
    }

    @Test(dependsOnMethods = {"buildAbbreviatedConjugations"})
    public void buildDetailConjugations() {
        ChartConfiguration chartConfiguration = getChartConfiguration().omitAbbreviatedConjugation(true);
        final ConjugationTemplate conjugationTemplate = getConjugationTemplate(chartConfiguration);
        MorphologicalChartEngine engine = new MorphologicalChartEngine(conjugationTemplate);
        List<MorphologicalChart> charts = engine.createMorphologicalCharts();
        final MorphologicalChart chart = charts.get(0);
        assertNotNull(chart);
        assertNull(chart.getAbbreviatedConjugation());

        final Path path = get(parentDocDir.toString(), "detail-conjugations.docx");

        DetailedConjugation[] detailedConjugations = new DetailedConjugation[charts.size()];
        for (int i = 0; i < charts.size(); i++) {
            detailedConjugations[i] = charts.get(i).getDetailedConjugation();
        }
        DetailedConjugationAdapter dca = new DetailedConjugationAdapter(detailedConjugations);
        try {
            WmlHelper.createDocument(path, chartConfiguration, dca);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        openFile(path);
    }

    private ChartConfiguration getChartConfiguration() {
        ChartConfiguration chartConfiguration = new ChartConfiguration();
        chartConfiguration.setArabicFontFamily(FontUtilities.getDefaultArabicFontName());
        chartConfiguration.setArabicFontSize(FontUtilities.DEFAULT_ARABIC_FONT_SIZE);
        chartConfiguration.setHeadingFontSize(FontUtilities.DEFAULT_ARABIC_FONT_SIZE);
        chartConfiguration.setTranslationFontFamily(FontUtilities.getDefaultEnglishFont());
        chartConfiguration.setTranslationFontSize(FontUtilities.DEFAULT_ENGLISH_FONT_SIZE);
        chartConfiguration.setHeadingFontSize(36L);
        return chartConfiguration;
    }

    private ConjugationTemplate getConjugationTemplate(ChartConfiguration chartConfiguration) {
        ConjugationTemplate conjugationTemplate = new ConjugationTemplate();
        conjugationTemplate.setChartConfiguration(chartConfiguration);
        conjugationTemplate.withData(getConjugationData(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, "To Help",
                new RootLetters(NOON, SAD, RA), VERBAL_NOUN_V1));
        conjugationTemplate.withData(getConjugationData(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, "To Say",
                new RootLetters(QAF, WAW, LAM), VERBAL_NOUN_V1));
        conjugationTemplate.withData(getConjugationData(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, "To Eat",
                new RootLetters(HAMZA, KAF, LAM), VERBAL_NOUN_V1));
        conjugationTemplate.withData(getConjugationData(FORM_IV_TEMPLATE, "To submit", new RootLetters(SEEN, LAM, MEEM)));
        conjugationTemplate.withData(getConjugationData(FORM_IV_TEMPLATE, "To send down", new RootLetters(NOON, ZAIN, LAM)));
        conjugationTemplate.withData(getConjugationData(FORM_IV_TEMPLATE, "To Establish", new RootLetters(QAF, WAW, MEEM)));
        conjugationTemplate.withData(getConjugationData(FORM_IX_TEMPLATE, "To collapse", new RootLetters(NOON, QAF, DDAD)));
        conjugationTemplate.withData(getConjugationData(FORM_VII_TEMPLATE, null, new RootLetters(KAF, SEEN, RA)));
        conjugationTemplate.withData(getConjugationData(FORM_VIII_TEMPLATE, null, new RootLetters(HAMZA, KHA, THAL)));
        conjugationTemplate.withData(getConjugationData(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, null, new RootLetters(MEEM, DAL, DAL)));
        return conjugationTemplate;
    }

    private ConjugationData getConjugationData(NamedTemplate template, String translation, RootLetters rootLetters,
                                               VerbalNoun... verbalNouns) {
        final ConjugationData conjugationData = new ConjugationData();
        conjugationData.setTemplate(template);
        conjugationData.setTranslation(translation);
        if (!ArrayUtils.isEmpty(verbalNouns)) {
            conjugationData.setVerbalNouns(Arrays.asList(verbalNouns));
        }
        conjugationData.setRootLetters(rootLetters);
        return conjugationData;
    }

}
