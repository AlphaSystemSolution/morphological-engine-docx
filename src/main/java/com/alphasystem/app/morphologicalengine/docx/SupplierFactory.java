package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalanalysis.morphology.model.ConjugationData;

/**
 * @author sali
 */
public interface SupplierFactory {

    MorphologicalChartSupplier createSupplier(ConjugationData conjugationData);
}
