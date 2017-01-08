package com.alphasystem.app.morphologicalengine.docx;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

/**
 * @author sali
 */
abstract class DocumentAdapter {

    protected abstract void buildDocument(MainDocumentPart mdp);
}
