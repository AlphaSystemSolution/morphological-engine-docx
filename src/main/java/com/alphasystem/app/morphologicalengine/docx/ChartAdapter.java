package com.alphasystem.app.morphologicalengine.docx;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Tbl;

/**
 * @author sali
 */
public abstract class ChartAdapter extends DocumentAdapter {

    @Override
    protected void buildDocument(MainDocumentPart mdp) {
        mdp.getContent().add(getChart());
    }

    protected abstract Tbl getChart();
}
