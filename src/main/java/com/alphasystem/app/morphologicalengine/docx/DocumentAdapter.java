package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 * @author sali
 */
abstract class DocumentAdapter {

    public ChartConfiguration getChartConfiguration() {
        return null;
    }

    protected abstract void buildDocument(MainDocumentPart mdp);
}
