package com.alphasystem.app.morphologicalengine.docx.test;

import java.awt.Desktop;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.lang3.ArrayUtils;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.testng.AbstractTestNGSpringContextTests;
import org.testng.Assert;
import org.testng.annotations.Test;

import com.alphasystem.app.morphologicalengine.docx.AbbreviatedConjugationAdapter;
import com.alphasystem.app.morphologicalengine.docx.AbbreviatedConjugationFactory;
import com.alphasystem.app.morphologicalengine.docx.DetailedConjugationAdapter;
import com.alphasystem.app.morphologicalengine.docx.DetailedConjugationFactory;
import com.alphasystem.app.morphologicalengine.docx.MorphologicalChartConfiguration;
import com.alphasystem.app.morphologicalengine.docx.MorphologicalChartEngine;
import com.alphasystem.app.morphologicalengine.docx.MorphologicalChartEngineFactory;
import com.alphasystem.app.morphologicalengine.docx.WmlHelper;
import com.alphasystem.app.morphologicalengine.spring.MorphologicalEngineConfiguration;
import com.alphasystem.arabic.model.NamedTemplate;
import com.alphasystem.arabic.ui.util.FontUtilities;
import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationData;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationTemplate;
import com.alphasystem.morphologicalanalysis.morphology.model.RootLetters;
import com.alphasystem.morphologicalanalysis.morphology.model.support.VerbalNoun;
import com.alphasystem.morphologicalengine.model.AbbreviatedConjugation;
import com.alphasystem.morphologicalengine.model.DetailedConjugation;
import com.alphasystem.morphologicalengine.model.MorphologicalChart;

import static com.alphasystem.arabic.model.ArabicLetterType.*;
import static com.alphasystem.arabic.model.NamedTemplate.*;
import static com.alphasystem.morphologicalanalysis.morphology.model.support.VerbalNoun.VERBAL_NOUN_V1;
import static java.lang.String.format;
import static java.lang.System.getProperty;
import static java.nio.file.Paths.get;
import static org.testng.Assert.assertNotNull;
import static org.testng.Assert.fail;
import static org.testng.Reporter.log;

/**
 * @author sali
 */
@ContextConfiguration(classes = {MorphologicalEngineConfiguration.class, MorphologicalChartConfiguration.class})
public class MorphologicalChartEngineTest extends AbstractTestNGSpringContextTests {

    private static final Path parentDocDir;

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

    @Autowired
    private MorphologicalChartEngineFactory morphologicalChartEngineFactory;
    @Autowired
    private AbbreviatedConjugationFactory abbreviatedConjugationFactory;
    @Autowired
    private DetailedConjugationFactory detailedConjugationFactory;

    @Test
    public void testCreateEmptyDocument() {
        final Path path = get(parentDocDir.toString(), "mydoc.docx");
        log(format("File Path: %s", path), true);
        MorphologicalChartEngine morphologicalChartEngine = morphologicalChartEngineFactory.createMorphologicalChartEngine(null);
        try {
            morphologicalChartEngine.createDocument(path);
            Assert.assertTrue(Files.exists(path));
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
    }

    @Test(dependsOnMethods = {"testCreateEmptyDocument"})
    public void runConjugationBuilder() {
        final Path path = get(parentDocDir.toString(), "conjugations.docx");
        final ConjugationTemplate conjugationTemplate = getConjugationTemplate(getChartConfiguration());
        MorphologicalChartEngine morphologicalChartEngine = morphologicalChartEngineFactory.createMorphologicalChartEngine(conjugationTemplate);
        try {
            morphologicalChartEngine.createDocument(path);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        openFile(path);
    }

    @Test(dependsOnMethods = {"runConjugationBuilder"})
    public void buildAbbreviatedConjugations() {
        final ChartConfiguration chartConfiguration = getChartConfiguration().omitToc(true).omitDetailedConjugation(true);
        final ConjugationTemplate conjugationTemplate = getConjugationTemplate(chartConfiguration);
        MorphologicalChartEngine engine = morphologicalChartEngineFactory.createMorphologicalChartEngine(conjugationTemplate);
        List<MorphologicalChart> charts = engine.createMorphologicalCharts();
        final MorphologicalChart chart = charts.get(0);
        assertNotNull(chart);

        final Path path = get(parentDocDir.toString(), "abbreviated-conjugations.docx");

        AbbreviatedConjugation[] abbreviatedConjugations = new AbbreviatedConjugation[charts.size()];
        for (int i = 0; i < charts.size(); i++) {
            abbreviatedConjugations[i] = charts.get(i).getAbbreviatedConjugation();
        }
        AbbreviatedConjugationAdapter aca = abbreviatedConjugationFactory.createAbbreviatedConjugationAdapter(chartConfiguration,
                abbreviatedConjugations);
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
        MorphologicalChartEngine engine = morphologicalChartEngineFactory.createMorphologicalChartEngine(conjugationTemplate);
        List<MorphologicalChart> charts = engine.createMorphologicalCharts();
        final MorphologicalChart chart = charts.get(0);
        assertNotNull(chart);

        final Path path = get(parentDocDir.toString(), "detail-conjugations.docx");

        DetailedConjugation[] detailedConjugations = new DetailedConjugation[charts.size()];
        for (int i = 0; i < charts.size(); i++) {
            detailedConjugations[i] = charts.get(i).getDetailedConjugation();
        }
        DetailedConjugationAdapter dca = detailedConjugationFactory.createDetailedConjugationAdapter(detailedConjugations);
        try {
            WmlHelper.createDocument(path, chartConfiguration, dca);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        openFile(path);
    }

    private ChartConfiguration getChartConfiguration() {
        ChartConfiguration chartConfiguration = new ChartConfiguration();
        chartConfiguration.setArabicFontFamily(FontUtilities.defaultArabicFontName);
        chartConfiguration.setArabicFontSize(FontUtilities.defaultArabicRegularFontSize);
        chartConfiguration.setHeadingFontSize(FontUtilities.defaultArabicHeadingFontSize);
        chartConfiguration.setTranslationFontFamily(FontUtilities.defaultEnglishFontName);
        chartConfiguration.setTranslationFontSize(FontUtilities.DEFAULT_ENGLISH_FONT_SIZE);
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
        conjugationTemplate.withData(getConjugationData(FORM_II_TEMPLATE, "To know", new RootLetters(AIN, LAM, MEEM)));
        conjugationTemplate.withData(getConjugationData(FORM_III_TEMPLATE, "To struggle", new RootLetters(JEEM, HA, DAL)));
        conjugationTemplate.withData(getConjugationData(FORM_IV_TEMPLATE, "To submit", new RootLetters(SEEN, LAM, MEEM)));
        conjugationTemplate.withData(getConjugationData(FORM_IV_TEMPLATE, "To send down", new RootLetters(NOON, ZAIN, LAM)));
        conjugationTemplate.withData(getConjugationData(FORM_IV_TEMPLATE, "To Establish", new RootLetters(QAF, WAW, MEEM)));
        conjugationTemplate.withData(getConjugationData(FORM_IX_TEMPLATE, "To collapse", new RootLetters(NOON, QAF, DDAD)));
        conjugationTemplate.withData(getConjugationData(FORM_VII_TEMPLATE, null, new RootLetters(KAF, SEEN, RA)));
        conjugationTemplate.withData(getConjugationData(FORM_VIII_TEMPLATE, null, new RootLetters(HAMZA, KHA, THAL)));
        conjugationTemplate.withData(getConjugationData(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, null, new RootLetters(MEEM, DAL, DAL)));
        conjugationTemplate.withData(getConjugationData(FORM_I_CATEGORY_A_GROUP_I_TEMPLATE, null, new RootLetters(DTHA, LAM, LAM)));
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
