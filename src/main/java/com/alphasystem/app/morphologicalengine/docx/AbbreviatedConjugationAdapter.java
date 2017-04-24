package com.alphasystem.app.morphologicalengine.docx;

import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.RPr;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;

import com.alphasystem.morphologicalanalysis.morphology.model.ChartConfiguration;
import com.alphasystem.morphologicalengine.model.AbbreviatedConjugation;
import com.alphasystem.morphologicalengine.model.ConjugationHeader;
import com.alphasystem.openxml.builder.wml.PBuilder;
import com.alphasystem.openxml.builder.wml.WmlAdapter;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;

import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.ADVERB_PREFIX;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.ARABIC_HEADING_STYLE;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.ARABIC_NORMAL_STYLE;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.COMMAND_PREFIX;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.FORBIDDING_PREFIX;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.PARTICIPLE_PREFIX;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.addSeparatorRow;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.getArabicTextP;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.getMultiWord;
import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.getNilBorderColumnProperties;
import static com.alphasystem.openxml.builder.wml.WmlAdapter.getText;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.BOOLEAN_DEFAULT_TRUE_TRUE;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.JC_CENTER;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getPBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getPPrBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getParaRPrBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRFontsBuilder;
import static com.alphasystem.openxml.builder.wml.WmlBuilderFactory.getRPrBuilder;
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

    AbbreviatedConjugationAdapter(ChartConfiguration chartConfiguration, AbbreviatedConjugation... abbreviatedConjugations) {
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
        String title = conjugationHeader.getTitle();
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

    private P getTranslationPara(String rsidR, String rsidP, String translation) {
        translation = (translation == null) ? "" : format("%s", translation);
        Text text = getText(translation, null);
        final String translationFontFamily = chartConfiguration.getTranslationFontFamily();
        RFonts rFonts = getRFontsBuilder().withAscii(translationFontFamily).withHAnsi(translationFontFamily).getObject();
        final long translationFontSize = chartConfiguration.getTranslationFontSize() * 2;
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withSz(translationFontSize).withSzCs(translationFontSize).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text).getObject();
        String rsidRpr = nextId();
        ParaRPr prpr = getParaRPrBuilder().withRFonts(rFonts).getObject();
        PPr ppr = getPPrBuilder().withJc(JC_CENTER).withRPr(prpr).getObject();
        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR).withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }

    private P getHeaderLabelPara(String rsidR, String rsidRpr, String rsidP, String label) {
        final long arabicFontSize = chartConfiguration.getArabicFontSize() * 2;
        ParaRPr prpr = getParaRPrBuilder().withSz(arabicFontSize).withSzCs(arabicFontSize).getObject();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_NORMAL_STYLE).withBidi(BOOLEAN_DEFAULT_TRUE_TRUE).withRPr(prpr).getObject();

        Text text = getText(label, null);
        final RFonts rFonts = getRFontsBuilder().withHint(CS).getObject();
        RPr rpr = getRPrBuilder().withRFonts(rFonts).withSz(arabicFontSize).withSzCs(arabicFontSize).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text).getObject();

        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR).withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }

    private void addActiveLineRow(AbbreviatedConjugation abbreviatedConjugation) {
        tableAdapter
                .startRow()
                .addColumn(0, getArabicTextP(PARTICIPLE_PREFIX, abbreviatedConjugation.getActiveParticipleMasculine()))
                .addColumn(1, getArabicTextP(getMultiWord(abbreviatedConjugation.getVerbalNouns())))
                .addColumn(2, getArabicTextP(abbreviatedConjugation.getPresentTense()))
                .addColumn(3, getArabicTextP(abbreviatedConjugation.getPastTense()))
                .endRow();
    }

    private void addPassiveLine(AbbreviatedConjugation abbreviatedConjugation) {
        tableAdapter
                .startRow()
                .addColumn(0, getArabicTextP(PARTICIPLE_PREFIX, abbreviatedConjugation.getPassiveParticipleMasculine()))
                .addColumn(1, getArabicTextP(getMultiWord(abbreviatedConjugation.getVerbalNouns())))
                .addColumn(2, getArabicTextP(abbreviatedConjugation.getPresentPassiveTense()))
                .addColumn(3, getArabicTextP(abbreviatedConjugation.getPastPassiveTense()))
                .endRow();
    }

    private void addCommandLine(AbbreviatedConjugation abbreviatedConjugation) {
        tableAdapter
                .startRow()
                .addColumn(0, 2, null, getArabicTextP(FORBIDDING_PREFIX,
                        abbreviatedConjugation.getForbidding()))
                .addColumn(2, 2, null, getArabicTextP(COMMAND_PREFIX,
                        abbreviatedConjugation.getImperative())).endRow();
    }

    private void addAdverbLine(AbbreviatedConjugation abbreviatedConjugation) {
        tableAdapter
                .startRow()
                .addColumn(0, 4, null, getArabicTextP(ADVERB_PREFIX, getMultiWord(abbreviatedConjugation.getAdverbs())))
                .endRow();
    }

}
