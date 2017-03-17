package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.builder.ConjugationBuilder;
import com.alphasystem.app.morphologicalengine.conjugation.builder.ConjugationHelper;
import com.alphasystem.app.morphologicalengine.conjugation.builder.ConjugationRoots;
import com.alphasystem.app.morphologicalengine.spring.MorphologicalEngineFactory;
import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationData;
import com.alphasystem.morphologicalanalysis.morphology.model.RootLetters;
import com.alphasystem.morphologicalengine.model.MorphologicalChart;

import java.util.function.Supplier;

/**
 * @author sali
 */
public class MorphologicalChartSupplier implements Supplier<MorphologicalChart> {

    private final ConjugationBuilder conjugationBuilder;
    private final ConjugationData conjugationData;

    public MorphologicalChartSupplier(ConjugationData conjugationData) {
        this.conjugationData = conjugationData;
        this.conjugationBuilder = MorphologicalEngineFactory.getConjugationBuilder();
    }

    private MorphologicalChart createChart() {
        if (conjugationData == null) {
            return null;
        }
        final RootLetters rootLetters = conjugationData.getRootLetters();
        if (rootLetters == null) {
            return null;
        }
        final ConjugationRoots conjugationRoots = ConjugationHelper.getConjugationRoots(conjugationData);
        return conjugationBuilder.doConjugation(conjugationRoots);
    }

    @Override
    public MorphologicalChart get() {
        return createChart();
    }

    @Override
    public String toString() {
        String result = super.toString();
        if (conjugationData == null) {
            return result;
        }
        return String.format("%s:%s", conjugationData.getTemplate(), conjugationData.getRootLetters().getDisplayName());
    }
}
