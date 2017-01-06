package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.model.AbbreviatedConjugation;
import com.alphasystem.app.morphologicalengine.conjugation.model.ConjugationHeader;
import com.alphasystem.app.morphologicalengine.conjugation.model.RootLetters;
import com.alphasystem.app.morphologicalengine.conjugation.model.abbrvconj.ActiveLine;
import com.alphasystem.app.morphologicalengine.conjugation.model.abbrvconj.AdverbLine;
import com.alphasystem.app.morphologicalengine.conjugation.model.abbrvconj.ImperativeAndForbiddingLine;
import com.alphasystem.app.morphologicalengine.conjugation.model.abbrvconj.PassiveLine;
import com.alphasystem.arabic.model.ArabicLetterType;
import com.alphasystem.arabic.model.ArabicWord;
import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.openxml.builder.wml.PBuilder;
import com.alphasystem.openxml.builder.wml.WmlAdapter;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import org.apache.commons.lang3.ArrayUtils;
import org.docx4j.wml.*;

import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.*;
import static com.alphasystem.arabic.model.ArabicWord.concatenateWithSpace;
import static com.alphasystem.fx.ui.util.FontConstants.ENGLISH_FONT_NAME;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getText;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.docx4j.wml.STHint.CS;

/**
 * @author sali
 */
public final class AbbreviatedConjugationAdapter extends ChartAdapter {

    private final ChartConfiguration chartConfiguration;
    private final AbbreviatedConjugation[] abbreviatedConjugations;
    private final TableAdapter tableAdapter;

    public AbbreviatedConjugationAdapter(ChartConfiguration chartConfiguration, AbbreviatedConjugation... abbreviatedConjugations) {
        this.chartConfiguration = (chartConfiguration == null) ? new ChartConfiguration() : chartConfiguration;
        this.abbreviatedConjugations = isEmpty(abbreviatedConjugations) ? new AbbreviatedConjugation[0] : abbreviatedConjugations;
        this.tableAdapter = new TableAdapter().startTable(25.0, 25.0, 25.0, 25.0);
    }

    @Override
    protected Tbl getChart() {
        for (AbbreviatedConjugation abbreviatedConjugation : abbreviatedConjugations) {
            if (!chartConfiguration.isOmitTitle()) {
                addTitleRow(abbreviatedConjugation);
            }
            if (!chartConfiguration.isOmitHeader()) {
                addHeaderRow(abbreviatedConjugation.getConjugationHeader());
            }
            addActiveLineRow(abbreviatedConjugation);
            addPassiveLine(abbreviatedConjugation);
            addCommandLine(abbreviatedConjugation);
            addAdverbLine(abbreviatedConjugation);
            addSeparatorRow(tableAdapter, 4);
        }

        return tableAdapter.getTable();
    }

    private void addTitleRow(AbbreviatedConjugation abbreviatedConjugation) {
        tableAdapter.startRow().addColumn(0, 4, getNilBorderColumnProperties(), createTitlePara(abbreviatedConjugation)).endRow();
    }

    private P createTitlePara(AbbreviatedConjugation abbreviatedConjugation) {
        final String id = nextId();

        final RFonts rFonts = getRFontsBuilder().withHint(CS).getObject();
        final RPr rpr = getRPrBuilder().withRFonts(rFonts).withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        final ConjugationHeader conjugationHeader = abbreviatedConjugation.getConjugationHeader();
        ArabicWord titleWord = (conjugationHeader == null) ? null : conjugationHeader.getTitle();
        titleWord = titleWord == null ? getTitleWord(abbreviatedConjugation.getActiveLine()) : titleWord;
        final String title = (titleWord == null) ? "" : titleWord.toUnicode();
        final Text text = getText(title);
        final R r = getRBuilder().withRsidRPr(id).withRPr(rpr).addContent(text).getObject();

        final ParaRPr paraRPr = getParaRPrBuilder().getObject();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_HEADING_STYLE).withBidi(BOOLEAN_DEFAULT_TRUE_TRUE).withRPr(paraRPr).getObject();

        final PBuilder pBuilder = getPBuilder().withParaId(id).withRsidP(id).withRsidR(id).withRsidRDefault(id).withRsidRPr(id)
                .withPPr(ppr).addContent(r);
        final P p = pBuilder.getObject();
        WmlAdapter.addBookMark(p, id);
        return p;
    }

    private void addHeaderRow(ConjugationHeader conjugationHeader) {
        String rsidR = nextId();
        String rsidP = nextId();

        // translation
        P translationPara = getTranslationPara(rsidR, rsidP, conjugationHeader.getTranslation());

        // second column paras
        String rsidRpr = nextId();
        P labelP1 = getHeaderLabelPara(rsidR, rsidRpr, rsidP, conjugationHeader.getTypeLabel1());
        P labelP2 = getHeaderLabelPara(rsidR, rsidRpr, rsidP, conjugationHeader.getTypeLabel2());
        P labelP3 = getHeaderLabelPara(rsidR, rsidRpr, rsidP, conjugationHeader.getTypeLabel3());

        tableAdapter.startRow()
                .addColumn(0, 2, null, WmlAdapter.getEmptyPara(), translationPara)
                .addColumn(2, 2, null, labelP1, labelP2, labelP3).endRow();
    }

    private P getRootWordsPara(String rsidR, String rsidP, ConjugationHeader conjugationHeader) {
        ParaRPr prpr = getParaRPrBuilder().withSz(SIZE_56).withSzCs(SIZE_56).getObject();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_NORMAL_STYLE).withBidi(BOOLEAN_DEFAULT_TRUE_TRUE).withJc(JC_CENTER)
                .withRPr(prpr).getObject();

        Text text = getText(getRootLetters(conjugationHeader));
        final RFonts rFonts = getRFontsBuilder().withHint(CS).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withSz(SIZE_56).withSzCs(SIZE_56).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text)
                .getObject();

        String rsidRpr = nextId();
        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR).withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }

    private String getRootLetters(ConjugationHeader conjugationHeader) {
        String result = "";
        if (conjugationHeader != null) {
            RootLetters rootLetters = conjugationHeader.getRootLetters();
            if (rootLetters != null) {
                ArabicWord[] rl = new ArabicWord[3];
                rl[0] = rootLetters.getFirstRadical().toLabel();
                rl[1] = rootLetters.getSecondRadical().toLabel();
                rl[2] = rootLetters.getThirdRadical().toLabel();
                final ArabicLetterType fourthRadical = rootLetters.getFourthRadical();
                if (fourthRadical != null) {
                    rl = ArrayUtils.add(rl, fourthRadical.toLabel());
                }
                result = concatenateWithSpace(rl).toUnicode();
            }
        }
        return result;
    }

    private P getTranslationPara(String rsidR, String rsidP, String translation) {
        translation = (translation == null) ? "" : format("%s", translation);
        Text text = getText(translation, null);
        RFonts rFonts = getRFontsBuilder().withAscii(ENGLISH_FONT_NAME).withHAnsi(ENGLISH_FONT_NAME).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text).getObject();
        String rsidRpr = nextId();
        ParaRPr prpr = getParaRPrBuilder().withRFonts(rFonts).getObject();
        PPr ppr = getPPrBuilder().withJc(JC_CENTER).withRPr(prpr).getObject();
        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR).withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }

    private P getHeaderLabelPara(String rsidR, String rsidRpr, String rsidP,
                                 ArabicWord label) {
        ParaRPr prpr = getParaRPrBuilder().withSz(SIZE_32).withSzCs(SIZE_32).getObject();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_NORMAL_STYLE).withBidi(BOOLEAN_DEFAULT_TRUE_TRUE).withRPr(prpr).getObject();

        Text text = getText(label.toUnicode(), null);
        final RFonts rFonts = getRFontsBuilder().withHint(CS).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withSz(SIZE_32).withSzCs(SIZE_32).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text).getObject();

        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR).withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }

    private void addActiveLineRow(AbbreviatedConjugation abbreviatedConjugation) {
        ActiveLine activeLine = abbreviatedConjugation.getActiveLine();
        if (activeLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, getArabicTextP(PARTICIPLE_PREFIX, activeLine.getActiveParticipleMasculine()))
                    .addColumn(1, getArabicTextP(getMultiWord(activeLine.getVerbalNouns())))
                    .addColumn(2, getArabicTextP(activeLine.getPresentTense()))
                    .addColumn(3, getArabicTextP(activeLine.getPastTense()))
                    .endRow();
        }
    }

    private void addPassiveLine(AbbreviatedConjugation abbreviatedConjugation) {
        PassiveLine passiveLine = abbreviatedConjugation.getPassiveLine();
        if (passiveLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, getArabicTextP(PARTICIPLE_PREFIX, passiveLine.getPassiveParticipleMasculine()))
                    .addColumn(1, getArabicTextP(getMultiWord(passiveLine.getVerbalNouns())))
                    .addColumn(2, getArabicTextP(passiveLine.getPresentPassiveTense()))
                    .addColumn(3, getArabicTextP(passiveLine.getPastPassiveTense()))
                    .endRow();
        }
    }

    private void addCommandLine(AbbreviatedConjugation abbreviatedConjugation) {
        ImperativeAndForbiddingLine commandLine = abbreviatedConjugation.getImperativeAndForbiddingLine();
        if (commandLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, 2, null, getArabicTextP(FORBIDDING_PREFIX, commandLine.getForbidding()))
                    .addColumn(2, 2, null, getArabicTextP(COMMAND_PREFIX, commandLine.getImperative())).endRow();
        }
    }

    private void addAdverbLine(AbbreviatedConjugation abbreviatedConjugation) {
        AdverbLine adverbLine = abbreviatedConjugation.getAdverbLine();
        if (adverbLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, 4, null, getArabicTextP(ADVERB_PREFIX, getMultiWord(adverbLine.getAdverbs())))
                    .endRow();
        }
    }

}
