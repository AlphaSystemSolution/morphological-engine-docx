package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import java.nio.file.Path;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.save;

/**
 * @author sali
 */
public abstract class DocumentAdapter {

    public void createDocument(Path path) throws Docx4JException {
        final WordprocessingMLPackage wordMLPackage = new WmlPackageBuilder().styles("META-INF/custom_styles.xml").getPackage();
        buildDocument(wordMLPackage.getMainDocumentPart());
        save(path.toFile(), wordMLPackage);
    }

    protected abstract void buildDocument(MainDocumentPart mdp);
}
