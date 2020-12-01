package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalengine.model.DetailedConjugation;

/**
 * @author sali
 */
public interface DetailedConjugationFactory {

    DetailedConjugationAdapter createDetailedConjugationAdapter(DetailedConjugation... detailedConjugations);
}
