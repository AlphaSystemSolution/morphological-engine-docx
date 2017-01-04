package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.model.abbrvconj.ActiveLine;
import com.alphasystem.arabic.model.ArabicSupport;
import com.alphasystem.arabic.model.ArabicWord;
import com.alphasystem.morphologicalanalysis.morphology.model.RootWord;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import org.docx4j.wml.*;

import static com.alphasystem.arabic.model.ArabicLetterType.*;
import static com.alphasystem.arabic.model.ArabicLetters.WORD_SPACE;
import static com.alphasystem.arabic.model.ArabicWord.*;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getNilBorders;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getText;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static org.apache.commons.lang3.ArrayUtils.isNotEmpty;
import static org.docx4j.wml.STHint.CS;

/**
 * @author sali
 */
final class WmlHelper {

    static final String ARABIC_HEADING_STYLE = "Arabic-Heading1";
    static final String ARABIC_NORMAL_STYLE = "Arabic-Normal";
    static final String ARABIC_CAPTION_STYLE = "Arabic-Caption";
    private static final String ARABIC_TABLE_CENTER_STYLE = "Arabic-Table-Center";
    static final String TRANSLATION_STYLE = "Georgia";
    private static final String NO_SPACING_STYLE = "NoSpacing";
    static final ArabicWord COMMAND_PREFIX = getWord(ALIF, LAM, ALIF_HAMZA_ABOVE, MEEM, RA, SPACE, MEEM, NOON, HA);
    static final ArabicWord FORBIDDING_PREFIX = getWord(WAW, NOON, HA, YA, SPACE, AIN, NOON, HA);
    static final ArabicWord ADVERB_PREFIX = getWord(WAW, ALIF, LAM, DTHA, RA, FA, SPACE, MEEM, NOON, HA);
    static final Long SIZE_56 = 56L;
    static final Long SIZE_32 = 32L;

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
        return getArabicTextP(null, value, ARABIC_TABLE_CENTER_STYLE);
    }

    static P getArabicTextP(ArabicWord prefix, ArabicSupport value) {
        return getArabicTextP(prefix, value, ARABIC_TABLE_CENTER_STYLE);
    }

    static P getArabicTextP(ArabicSupport value, String pStyle) {
        return getArabicTextP(null, value, pStyle);
    }

    private static P getArabicTextP(ArabicWord prefix, ArabicSupport value, String pStyle) {
        String rsidr = nextId();
        PPr ppr = getPPrBuilder().withPStyle(pStyle).getObject();
        final RFonts rFonts = getRFontsBuilder().withHint(CS).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        ArabicWord word = (value == null) ? WORD_SPACE : value.toLabel();
        if (prefix != null) {
            word = concatenateWithSpace(prefix, word);
        }
        Text text = getText(word.toUnicode(), null);
        String id = nextId();
        R r = getRBuilder().withRsidRPr(id).withRPr(rpr).addContent(text).getObject();
        return getPBuilder().withRsidR(rsidr).withRsidRDefault(rsidr).withRsidRPr(id).withRsidP(id).withPPr(ppr).addContent(r)
                .getObject();
    }

    static TcPr getNilBorderColumnProperties() {
        return getTcPrBuilder().withTcBorders(getNilBorders()).getObject();
    }
}
