package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalengine.model.AbbreviatedConjugation;

/**
 * @author sali
 */
public interface AbbreviatedConjugationFactory {

    AbbreviatedConjugationAdapter createAbbreviatedConjugationAdapter(ChartConfiguration chartConfiguration,
                                                                      AbbreviatedConjugation... abbreviatedConjugations);
}
