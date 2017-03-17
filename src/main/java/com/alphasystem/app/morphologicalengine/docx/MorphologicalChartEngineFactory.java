package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationTemplate;

/**
 * @author sali
 */
public interface MorphologicalChartEngineFactory {

    MorphologicalChartEngine createMorphologicalChartEngine(ConjugationTemplate conjugationTemplate);
}
