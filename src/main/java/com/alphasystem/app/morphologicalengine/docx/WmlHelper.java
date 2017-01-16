package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.model.abbrvconj.ActiveLine;
import com.alphasystem.arabic.model.ArabicSupport;
import com.alphasystem.arabic.model.ArabicWord;
import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalanalysis.morphology.model.RootWord;
import com.alphasystem.morphologicalanalysis.morphology.model.support.PageOrientation;
import com.alphasystem.openxml.builder.wml.PPrBuilder;
import com.alphasystem.openxml.builder.wml.StylesBuilder;
import com.alphasystem.openxml.builder.wml.WmlBuilderFactory;
import com.alphasystem.openxml.builder.wml.WmlPackageBuilder;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import com.alphasystem.util.IdGenerator;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.wml.*;

import java.nio.file.Path;

import static com.alphasystem.arabic.model.ArabicLetterType.*;
import static com.alphasystem.arabic.model.ArabicLetters.WORD_SPACE;
import static com.alphasystem.arabic.model.ArabicWord.concatenateWithAnd;
import static com.alphasystem.arabic.model.ArabicWord.concatenateWithSpace;
import static com.alphasystem.arabic.model.ArabicWord.getWord;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getNilBorders;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getText;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.save;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.BOOLEAN_DEFAULT_TRUE_TRUE;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getColorBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getPBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getPPrBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRFontsBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRPrBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getStyleBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getTcPrBuilder;
import static com.alphasystem.util.IdGenerator.nextId;
import static org.apache.commons.lang3.ArrayUtils.isNotEmpty;
import static org.docx4j.wml.JcEnumeration.CENTER;
import static org.docx4j.wml.STHint.CS;
import static org.docx4j.wml.STThemeColor.ACCENT_1;

/**
 * @author sali
 */
public final class WmlHelper {

    static final String ARABIC_HEADING_STYLE = "Arabic-Heading1";
    static final String ARABIC_NORMAL_STYLE = "Arabic-Normal";
    static final String ARABIC_CAPTION_STYLE = "Arabic-Caption";
    private static final String ARABIC_TABLE_CENTER_STYLE = "Arabic-Table-Center";
    private static final String ARABIC_PREFIX_STYLE = "Arabic-PrefixChar";
    private static final String NO_SPACING_STYLE = "NoSpacing";
    static final ArabicWord PARTICIPLE_PREFIX = getWord(FA, HA, WAW);
    static final ArabicWord COMMAND_PREFIX = getWord(ALIF, LAM, ALIF_HAMZA_ABOVE, MEEM, RA, SPACE, MEEM, NOON, HA);
    static final ArabicWord FORBIDDING_PREFIX = getWord(WAW, NOON, HA, YA, SPACE, AIN, NOON, HA);
    static final ArabicWord ADVERB_PREFIX = getWord(WAW, ALIF, LAM, DTHA, RA, FA, SPACE, MEEM, NOON, HA);

    public static void createDocument(Path path, ChartConfiguration chartConfiguration, DocumentAdapter documentAdapter) throws Docx4JException {
        final String fontFamily = chartConfiguration.getArabicFontFamily();
        long normalFontSize = chartConfiguration.getArabicFontSize() * 2;
        long headingFontSize = chartConfiguration.getHeadingFontSize() * 2;
        final PageOrientation orientation = chartConfiguration.getPageOption().getOrientation();
        boolean landscape = PageOrientation.LANDSCAPE.equals(orientation);

        final WordprocessingMLPackage wordMLPackage = WmlPackageBuilder.createPackage(landscape).styles(
                createStyles(fontFamily, normalFontSize, headingFontSize)).getPackage();
        documentAdapter.buildDocument(wordMLPackage.getMainDocumentPart());
        save(path.toFile(), wordMLPackage);
    }

    static void addSeparatorRow(TableAdapter tableAdapter, Integer gridSpan) {
        final TcPr tcPr = getTcPrBuilder().withTcBorders(getNilBorders()).getObject();
        tableAdapter.startRow()
                .addColumn(0, gridSpan, tcPr, createNoSpacingStyleP())
                .endRow();
    }

    static P createNoSpacingStyleP() {
        PPr ppr = getPPrBuilder().withPStyle(NO_SPACING_STYLE).getObject();
        return getPBuilder().withRsidR(nextId()).withRsidP(nextId())
                .withRsidRDefault(nextId()).withPPr(ppr).getObject();
    }

    /**
     * Gets the title of the conjugation. The title will be comprised of third
     * person singular masculine past tense (space> third person singular
     * masculine present tense.
     *
     * @param activeLine active line
     * @return title {@link ArabicWord}
     */
    static ArabicWord getTitleWord(ActiveLine activeLine) {
        ArabicWord pastTense = WORD_SPACE;
        ArabicWord presentTense = WORD_SPACE;
        if (activeLine != null) {
            RootWord pastTenseRootWord = activeLine.getPastTense();
            pastTense = (pastTenseRootWord == null) ? WORD_SPACE : pastTenseRootWord.getRootWord();
            if (pastTense == null) {
                pastTense = WORD_SPACE;
            }
            RootWord presentTenseRootWord = activeLine.getPresentTense();
            presentTense = (presentTenseRootWord == null) ? WORD_SPACE : presentTenseRootWord.getRootWord();
            if (presentTense == null) {
                presentTense = WORD_SPACE;
            }
        }
        return concatenateWithSpace(pastTense, presentTense);
    }

    static ArabicWord getMultiWord(RootWord[] words) {
        ArabicWord w = WORD_SPACE;
        if (isNotEmpty(words)) {
            w = words[0].toLabel();
            for (int i = 1; i < words.length; i++) {
                w = concatenateWithAnd(w, words[i].toLabel());
            }
        }
        return w;
    }

    static P getArabicTextP(ArabicSupport value) {
        return getArabicTextP(value, ARABIC_TABLE_CENTER_STYLE);
    }

    private static P getArabicTextP(ArabicWord prefix, ArabicSupport value, String pStyle, String prefixStyle) {
        if (prefix == null) {
            return getArabicTextP(value, pStyle);
        }
        String rsidr = nextId();
        PPr ppr = getPPrBuilder().withPStyle(pStyle).getObject();


        RFonts rFonts = getRFontsBuilder().withHint(CS).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withRStyle(prefixStyle).withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        Text text = getText(prefix.toUnicode() + " ", "preserve");
        String id = nextId();
        R prefixRun = getRBuilder().withRsidRPr(id).withRPr(rpr).addContent(text).getObject();

        ArabicWord word = (value == null) ? WORD_SPACE : value.toLabel();
        text = getText(word.toUnicode(), null);
        id = nextId();
        rFonts = getRFontsBuilder().withHint(CS).getObject();
        rpr = getRPrBuilder().withRFonts(rFonts).withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        R mainRun = getRBuilder().withRsidRPr(id).withRPr(rpr).addContent(text).getObject();

        return getPBuilder().withRsidR(rsidr).withRsidRDefault(rsidr).withRsidRPr(id).withRsidP(id).withPPr(ppr)
                .addContent(prefixRun, mainRun)
                .getObject();
    }

    static P getArabicTextP(ArabicWord prefix, ArabicSupport value) {
        return getArabicTextP(prefix, value, ARABIC_TABLE_CENTER_STYLE, ARABIC_PREFIX_STYLE);
    }

    static P getArabicTextP(ArabicSupport value, String pStyle) {
        String rsidr = nextId();
        PPr ppr = getPPrBuilder().withPStyle(pStyle).getObject();
        final RFonts rFonts = getRFontsBuilder().withHint(CS).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        ArabicWord word = (value == null) ? WORD_SPACE : value.toLabel();
        Text text = getText(word.toUnicode(), null);
        String id = nextId();
        R r = getRBuilder().withRsidRPr(id).withRPr(rpr).addContent(text).getObject();
        return getPBuilder().withRsidR(rsidr).withRsidRDefault(rsidr).withRsidRPr(id).withRsidP(id).withPPr(ppr).addContent(r)
                .getObject();
    }

    static TcPr getNilBorderColumnProperties() {
        return getTcPrBuilder().withTcBorders(getNilBorders()).getObject();
    }

    private static Styles createStyles(String family, long normalSize, long headingSize) {
        final StylesBuilder stylesBuilder = WmlBuilderFactory.getStylesBuilder();
        return stylesBuilder.addStyle(createArabicNormalStyle(family, normalSize), createArabicNormalCharStyle(family, normalSize),
                createArabicTableCenterStyle(), createArabicTableCenterCharStyle(family, normalSize),
                createArabicTableCaptionStyle(), createArabicTableCaptionCharStyle(family, normalSize),
                createArabicHeading1Style(family, headingSize), createArabicHeading1CharStyle(family, headingSize),
                createArabicTocStyle(family), createArabicPrefixCharStyle(family, 24),
                createArabicPrefixParagraphStyle(family, 24)).getObject();
    }

    private static Style createArabicNormalStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        PPr ppr = getPPrBuilder().withBidi(true).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Normal").withCustomStyle(true)
                .withName("Arabic-Normal").withBasedOn("Normal").withLink("Arabic-NormalChar").withQFormat(true)
                .withRsid("0004111F").withPPr(ppr).withRPr(rpr).getObject();
    }

    private static Style createArabicNormalCharStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-NormalChar").withCustomStyle(true)
                .withName("Arabic-Normal Char").withBasedOn("DefaultParagraphFont").withLink("Arabic-Normal")
                .withRsid("0004111F").withRPr(rpr).getObject();
    }

    private static Style createArabicTableCenterStyle() {
        final PPrBuilder pPrBuilder = getPPrBuilder();
        final PPrBase.Spacing spacing = pPrBuilder.getSpacingBuilder().withBefore(120L).withAfter(120L).withLine(240L)
                .withLineRule(STLineSpacingRule.AUTO).getObject();
        PPr ppr = pPrBuilder.withJc(CENTER).withSpacing(spacing).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Table-Center").withCustomStyle(true)
                .withName("Arabic-Table-Center").withBasedOn("Arabic-Normal").withLink("Arabic-Normal").withQFormat(true)
                .withRsid("00B94679").withPPr(ppr).getObject();
    }

    private static Style createArabicTableCenterCharStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-Table-CenterChar").withCustomStyle(true)
                .withName("Arabic-Table-Center Char").withBasedOn("Arabic-NormalChar").withLink("Arabic-Table-Center")
                .withRsid("00B94679").withRPr(rpr).getObject();
    }

    private static Style createArabicTableCaptionStyle() {
        PPr ppr = getPPrBuilder().withBidi(true).getObject();
        Color color = getColorBuilder().withVal("2E74B5").withThemeColor(ACCENT_1).withThemeShade("BF").getObject();
        RPr rpr = getRPrBuilder().withB(true).withBCs(true).withColor(color).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Caption").withCustomStyle(true)
                .withName("Arabic-Caption").withBasedOn("Arabic-Table-Center").withLink("Arabic-CaptionChar").withQFormat(true)
                .withRsid("00141798").withPPr(ppr).withRPr(rpr).getObject();
    }

    private static Style createArabicTableCaptionCharStyle(String family, long size) {
        Color color = getColorBuilder().withVal("2E74B5").withThemeColor(ACCENT_1).withThemeShade("BF").getObject();
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).withB(true).withBCs(true).withColor(color)
                .getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-CaptionChar").withCustomStyle(true)
                .withName("Arabic-Caption Char").withBasedOn("Arabic-Table-CenterChar").withLink("Arabic-Caption")
                .withRsid("00141798").withRPr(rpr).getObject();
    }

    private static Style createArabicHeading1Style(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withRFonts(rFonts).getObject();
        PPr ppr = getPPrBuilder().withBidi(true).withJc(CENTER).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Heading1").withCustomStyle(true)
                .withName("Arabic-Heading1").withBasedOn("Heading1").withLink("Arabic-Heading1Char").withQFormat(true)
                .withRsid("002D29F2").withPPr(ppr).withRPr(rpr).getObject();
    }

    private static Style createArabicHeading1CharStyle(String family, long size) {
        Color color = getColorBuilder().withVal("2E74B5").withThemeColor(ACCENT_1).withThemeShade("BF").getObject();
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        RPr rpr = getRPrBuilder().withSz(size).withSzCs(size).withB(true).withBCs(true).withColor(color).withRFonts(rFonts).getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-Heading1Char").withCustomStyle(true)
                .withName("Arabic-Heading1 Char").withLink("Arabic-Heading1")
                .withRsid("002D29F2").withRPr(rpr).getObject();
    }

    private static Style createArabicTocStyle(String family) {
        CTTabStop tab = WmlBuilderFactory.getCTTabStopBuilder().withVal(STTabJc.RIGHT).withLeader(STTabTlc.DOT)
                .withPos(9017L).getObject();
        Tabs tabs = WmlBuilderFactory.getTabsBuilder().addTab(tab).getObject();
        PPr ppr = getPPrBuilder().withTabs(tabs).getObject();
        RFonts rFonts = getRFontsBuilder().withCs(family).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("TOCArabic").withCustomStyle(true)
                .withName("TOC Arabic").withBasedOn("TOC1").withQFormat(true).withRsid(IdGenerator.nextId())
                .withPPr(ppr).withRPr(rpr).getObject();
    }

    private static Style createArabicPrefixCharStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        Color color = getColorBuilder().withVal("C00000").getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withColor(color).withSz(size).withSzCs(size).getObject();
        return getStyleBuilder().withType("character").withStyleId("Arabic-PrefixChar").withCustomStyle(true)
                .withName("Arabic-Prefix Char").withLink("Arabic-Prefix").withBasedOn("Arabic-Table-CenterChar")
                .withRsid("005B3CB5").withRPr(rpr).getObject();
    }

    private static Style createArabicPrefixParagraphStyle(String family, long size) {
        RFonts rFonts = getRFontsBuilder().withAscii(family).withHAnsi(family).withCs(family).getObject();
        Color color = getColorBuilder().withVal("C00000").getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withColor(color).withSz(size).withSzCs(size).getObject();
        return getStyleBuilder().withType("paragraph").withStyleId("Arabic-Prefix").withCustomStyle(true)
                .withName("Arabic-Prefix").withLink("Arabic-PrefixChar").withBasedOn("Arabic-Table-Center")
                .withRsid("005B3CB5").withRPr(rpr).getObject();
    }
}
