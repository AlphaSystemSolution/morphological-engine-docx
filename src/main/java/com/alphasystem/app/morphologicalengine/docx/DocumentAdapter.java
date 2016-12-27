package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.openxml.builder.wml.PPrBuilder;
import com.alphasystem.openxml.builder.wml.StylesBuilder;
import com.alphasystem.openxml.builder.wml.WmlBuilderFactory;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import javafx.scene.text.Font;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.Spacing;

import java.nio.file.Path;
import java.util.List;

import static com.alphasystem.openxml.builder.wml.WmlAdapter.save;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static org.docx4j.wml.JcEnumeration.CENTER;
import static org.docx4j.wml.STThemeColor.ACCENT_1;

/**
 * @author sali
 */
public abstract class DocumentAdapter {

    public void createDocument(Path path) throws Docx4JException {
        final String fontFamily = getArabicFontFamily();
        long normalFontSize = Long.getLong("arabic.normal.font.size", 40L);
        long headingFontSize = Long.getLong("arabic.heading.font.size", 72L);

        final WordprocessingMLPackage wordMLPackage = new WmlPackageBuilder().styles(
                createStyles(fontFamily, normalFontSize, headingFontSize)).getPackage();
        buildDocument(wordMLPackage.getMainDocumentPart());
        save(path.toFile(), wordMLPackage);
    }

    private String getArabicFontFamily() {
        String fontFamily = System.getProperty("arabic.font.name");
        if (fontFamily == null) {
            fontFamily = "Traditional Arabic";
        }
        final List<String> families = Font.getFamilies();
        if (families.contains(fontFamily)) {
            return fontFamily;
        }
        fontFamily = "Traditional Arabic";
        if (families.contains(fontFamily)) {
            return fontFamily;
        }
        fontFamily = "Arabic Typesetting";
        if (families.contains(fontFamily)) {
            return fontFamily;
        }
        return "Arial";
    }

    private Styles createStyles(String family, long normalSize, long headingSize) {
        final StylesBuilder stylesBuilder = WmlBuilderFactory.getStylesBuilder();
        return stylesBuilder.addStyle(createArabicNormalStyle(family, normalSize), createArabicNormalCharStyle(family, normalSize),
                createArabicTableCenterStyle(), createArabicTableCenterCharStyle(family, normalSize),
                createArabicTableCaptionStyle(), createArabicTableCaptionCharStyle(family, normalSize),
                createArabicHeading1Style(family, headingSize), createArabicHeading1CharStyle(family, headingSize)).getObject();
    }

    private Style createArabicNormalStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        PPr ppr = getPPrBuilder().withBidi(true).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Normal").withCustomStyle(true)
                .withName("Arabic-Normal").withBasedOn("Normal").withLink("Arabic-NormalChar").withQFormat(true)
                .withRsid("0004111F").withPPr(ppr).withRPr(rpr).getObject();
    }

    private Style createArabicNormalCharStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-NormalChar").withCustomStyle(true)
                .withName("Arabic-Normal Char").withBasedOn("DefaultParagraphFont").withLink("Arabic-Normal")
                .withRsid("0004111F").withRPr(rpr).getObject();
    }

    private Style createArabicTableCenterStyle() {
        final PPrBuilder pPrBuilder = getPPrBuilder();
        final Spacing spacing = pPrBuilder.getSpacingBuilder().withBefore(120L).withAfter(120L).withLine(240L)
                .withLineRule(STLineSpacingRule.AUTO).getObject();
        PPr ppr = pPrBuilder.withJc(CENTER).withSpacing(spacing).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Table-Center").withCustomStyle(true)
                .withName("Arabic-Table-Center").withBasedOn("Arabic-Normal").withLink("Arabic-Normal").withQFormat(true)
                .withRsid("00B94679").withPPr(ppr).getObject();
    }

    private Style createArabicTableCenterCharStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-Table-CenterChar").withCustomStyle(true)
                .withName("Arabic-Table-Center Char").withBasedOn("Arabic-NormalChar").withLink("Arabic-Table-Center")
                .withRsid("00B94679").withRPr(rpr).getObject();
    }

    private Style createArabicTableCaptionStyle() {
        PPr ppr = getPPrBuilder().withBidi(true).getObject();
        Color color = getColorBuilder().withVal("2E74B5").withThemeColor(ACCENT_1).withThemeShade("BF").getObject();
        RPr rpr = getRPrBuilder().withB(true).withBCs(true).withColor(color).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Caption").withCustomStyle(true)
                .withName("Arabic-Caption").withBasedOn("Arabic-Table-Center").withLink("Arabic-CaptionChar").withQFormat(true)
                .withRsid("00141798").withPPr(ppr).withRPr(rpr).getObject();
    }

    private Style createArabicTableCaptionCharStyle(String family, long size) {
        Color color = getColorBuilder().withVal("2E74B5").withThemeColor(ACCENT_1).withThemeShade("BF").getObject();
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).withB(true).withBCs(true).withColor(color)
                .getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-CaptionChar").withCustomStyle(true)
                .withName("Arabic-Caption Char").withBasedOn("Arabic-Table-CenterChar").withLink("Arabic-Caption")
                .withRsid("00141798").withRPr(rpr).getObject();
    }

    private Style createArabicHeading1Style(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        PPr ppr = getPPrBuilder().withBidi(true).withJc(CENTER).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Heading1").withCustomStyle(true)
                .withName("Arabic-Heading1").withBasedOn("Heading1").withLink("Arabic-Heading1Char").withQFormat(true)
                .withRsid("002D29F2").withPPr(ppr).withRPr(rpr).getObject();
    }

    private Style createArabicHeading1CharStyle(String family, long size) {
        Color color = getColorBuilder().withVal("2E74B5").withThemeColor(ACCENT_1).withThemeShade("BF").getObject();
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withB(true).withBCs(true).withColor(color).withRFonts(rFonts).getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-Heading1Char").withCustomStyle(true)
                .withName("Arabic-Heading1 Char").withLink("Arabic-Heading1")
                .withRsid("002D29F2").withRPr(rpr).getObject();
    }

    protected abstract void buildDocument(MainDocumentPart mdp);
}
