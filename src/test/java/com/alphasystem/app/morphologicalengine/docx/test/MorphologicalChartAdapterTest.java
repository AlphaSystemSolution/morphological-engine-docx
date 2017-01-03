package com.alphasystem.app.morphologicalengine.docx.test;

import com.alphasystem.app.morphologicalengine.conjugation.builder.ConjugationBuilder;
import com.alphasystem.app.morphologicalengine.conjugation.builder.ConjugationRoots;
import com.alphasystem.app.morphologicalengine.conjugation.model.AbbreviatedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.DetailedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.MorphologicalChart;
import com.alphasystem.app.morphologicalengine.conjugation.model.NounRootBase;
import com.alphasystem.app.morphologicalengine.docx.AbbreviatedConjugationAdapter;
import com.alphasystem.app.morphologicalengine.docx.DetailedConjugationAdapter;
import com.alphasystem.app.morphologicalengine.docx.MorphologicalChartEngine;
import com.alphasystem.app.morphologicalengine.guice.GuiceSupport;
import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.testng.Assert;
import org.testng.annotations.Test;

import java.awt.*;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;

import static com.alphasystem.app.morphologicalengine.conjugation.builder.ConjugationHelper.getConjugationRoots;
import static com.alphasystem.arabic.model.ArabicLetterType.*;
import static com.alphasystem.arabic.model.NamedTemplate.*;
import static com.alphasystem.morphologicalanalysis.morphology.model.support.BrokenPlural.BROKEN_PLURAL_V12;
import static com.alphasystem.morphologicalanalysis.morphology.model.support.NounOfPlaceAndTime.*;
import static com.alphasystem.morphologicalanalysis.morphology.model.support.VerbalNoun.VERBAL_NOUN_V1;
import static java.lang.String.format;
import static java.lang.System.getProperty;
import static java.nio.file.Paths.get;
import static org.apache.commons.lang3.ArrayUtils.add;
import static org.testng.Assert.*;
import static org.testng.Reporter.log;

/**
 * @author sali
 */
public class MorphologicalChartAdapterTest {

    private static final NounRootBase[] FORM_I_ADVERBS = new NounRootBase[]{
            new NounRootBase(NOUN_OF_PLACE_AND_TIME_V1, BROKEN_PLURAL_V12),
            new NounRootBase(NOUN_OF_PLACE_AND_TIME_V2, BROKEN_PLURAL_V12),
            new NounRootBase(NOUN_OF_PLACE_AND_TIME_V3)};

    private static Path parentDocDir = null;

    static {
        parentDocDir = get(getProperty("target.dir"), "docs");
        try {
            Files.createDirectories(parentDocDir);
        } catch (IOException e) {
            e.printStackTrace();
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
        MorphologicalChart[] charts = getCharts(null);
        final Path path = get(parentDocDir.toString(), "conjugations.docx");
        MorphologicalChartEngine morphologicalChartEngine = new MorphologicalChartEngine(charts);
        try {
            morphologicalChartEngine.createDocument(path);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        try {
            Desktop.getDesktop().open(path.toFile());
        } catch (IOException e) {
            // ignore
        }
    }

    @Test(dependsOnMethods = {"runConjugationBuilder"})
    public void buildAbbreviatedConjugations() {
        ChartConfiguration chartConfiguration = new ChartConfiguration().omitToc(true).omitDetailedConjugation(true);
        MorphologicalChart[] charts = getCharts(chartConfiguration);
        final MorphologicalChart chart = charts[0];
        assertNotNull(chart);
        assertNull(chart.getDetailedConjugation());

        final Path path = get(parentDocDir.toString(), "abbreviated-conjugations.docx");

        AbbreviatedConjugation[] abbreviatedConjugations = new AbbreviatedConjugation[charts.length];
        for (int i = 0; i < charts.length; i++) {
            abbreviatedConjugations[i] = charts[i].getAbbreviatedConjugation();
        }
        AbbreviatedConjugationAdapter aca = new AbbreviatedConjugationAdapter(chartConfiguration, abbreviatedConjugations);
        try {
            aca.createDocument(path);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        try {
            Desktop.getDesktop().open(path.toFile());
        } catch (IOException e) {
            // ignore
        }
    }

    @Test(dependsOnMethods = {"buildAbbreviatedConjugations"})
    public void buildDetailConjugations() {
        ChartConfiguration chartConfiguration = new ChartConfiguration().omitAbbreviatedConjugation(true);
        MorphologicalChart[] charts = getCharts(chartConfiguration);
        final MorphologicalChart chart = charts[0];
        assertNotNull(chart);
        assertNull(chart.getAbbreviatedConjugation());

        final Path path = get(parentDocDir.toString(), "detail-conjugations.docx");

        DetailedConjugation[] detailedConjugations = new DetailedConjugation[charts.length];
        for (int i = 0; i < charts.length; i++) {
            detailedConjugations[i] = charts[i].getDetailedConjugation();
        }
        DetailedConjugationAdapter dca = new DetailedConjugationAdapter(detailedConjugations);
        try {
            dca.createDocument(path);
        } catch (Docx4JException e) {
            fail(format("Failed to create document {%s}", path), e);
        }
        try {
            Desktop.getDesktop().open(path.toFile());
        } catch (IOException e) {
            // ignore
        }
    }

    private MorphologicalChart[] getCharts(ChartConfiguration chartConfiguration) {
        MorphologicalChart[] charts = new MorphologicalChart[0];
        final ConjugationBuilder conjugationBuilder = GuiceSupport.getInstance().getConjugationBuilder();

        ConjugationRoots conjugationRoots = getConjugationRoots(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, "To Help",
                new NounRootBase[]{new NounRootBase(VERBAL_NOUN_V1)}, FORM_I_ADVERBS).chartConfiguration(chartConfiguration);
        charts = add(charts, conjugationBuilder.doConjugation(conjugationRoots, NOON, SAD, RA, null));

        conjugationRoots = getConjugationRoots(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, "To Say",
                new NounRootBase[]{new NounRootBase(VERBAL_NOUN_V1)}, FORM_I_ADVERBS).chartConfiguration(chartConfiguration);
        charts = add(charts, conjugationBuilder.doConjugation(conjugationRoots, QAF, WAW, LAM, null));

        conjugationRoots = getConjugationRoots(FORM_I_CATEGORY_A_GROUP_U_TEMPLATE, "To Eat",
                new NounRootBase[]{new NounRootBase(VERBAL_NOUN_V1)}, FORM_I_ADVERBS).chartConfiguration(chartConfiguration);
        charts = add(charts, conjugationBuilder.doConjugation(conjugationRoots, HAMZA, KAF, LAM, null));

        conjugationRoots = getConjugationRoots(FORM_IV_TEMPLATE, "To submit").chartConfiguration(chartConfiguration);
        charts = add(charts, (conjugationBuilder.doConjugation(conjugationRoots, SEEN, LAM, MEEM, null)));

        conjugationRoots = getConjugationRoots(FORM_IV_TEMPLATE, "To send down").chartConfiguration(chartConfiguration);
        charts = add(charts, (conjugationBuilder.doConjugation(conjugationRoots, NOON, ZAIN, LAM, null)));

        conjugationRoots = getConjugationRoots(FORM_IV_TEMPLATE, "To Establish").chartConfiguration(chartConfiguration);
        charts = add(charts, (conjugationBuilder.doConjugation(conjugationRoots, QAF, WAW, MEEM, null)));

        conjugationRoots = getConjugationRoots(FORM_IX_TEMPLATE, "To collapse").chartConfiguration(chartConfiguration);
        charts = add(charts, (conjugationBuilder.doConjugation(conjugationRoots, NOON, QAF, DDAD, null)));
        return charts;
    }

}
