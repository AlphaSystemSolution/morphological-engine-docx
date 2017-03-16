package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.app.morphologicalengine.conjugation.model.*;
import com.alphasystem.morphologicalanalysis.morphology.model.RootWord;
import com.alphasystem.morphologicalanalysis.morphology.model.support.SarfTermType;
import com.alphasystem.morphologicalengine.model.ConjugationTuple;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import com.alphasystem.openxml.builder.wml.table.TableAdapter.VerticalMergeType;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TcPr;

import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.*;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;

/**
 * @author sali
 */
public final class DetailedConjugationAdapter extends ChartAdapter {

    private static final int NUM_OF_COLUMNS = 7;

    private final DetailedConjugation[] detailedConjugations;
    private final TableAdapter tableAdapter;

    public DetailedConjugationAdapter(DetailedConjugation... detailedConjugations) {
        this.detailedConjugations = isEmpty(detailedConjugations) ? new DetailedConjugation[0] : detailedConjugations;
        this.tableAdapter = new TableAdapter().startTable(16.24, 16.24, 16.24, 2.56, 16.24, 16.24, 16.24);
    }

    private static TcPr getColumnProperties(Object obj) {
        return (obj == null) ? getNilBorderColumnProperties() : null;
    }

    @Override
    protected Tbl getChart() {
        for (DetailedConjugation detailedConjugation : detailedConjugations) {
            addTensePair(detailedConjugation.getActiveTensePair());
            addNounPair(detailedConjugation.getActiveParticiplePair());
            addNounPairs(detailedConjugation.getVerbalNounPairs());
            addTensePair(detailedConjugation.getPassiveTensePair());
            addNounPair(detailedConjugation.getPassiveParticiplePair());
            addTensePair(detailedConjugation.getImperativeAndForbiddingPair());
            addNounPairs(detailedConjugation.getAdverbPairs());
        }

        return tableAdapter.getTable();
    }

    private void addTensePair(VerbDetailedConjugationPair conjugationPair) {
        if (conjugationPair == null) {
            return;
        }
        final VerbConjugationGroup leftSideConjugations = conjugationPair.getLeftSideConjugations();
        final VerbConjugationGroup rightSideConjugations = conjugationPair.getRightSideConjugations();
        final boolean noLeftConjugations = leftSideConjugations == null;
        final boolean noRightConjugations = rightSideConjugations == null;

        final SarfTermType leftSideCaption = noLeftConjugations ? null : leftSideConjugations.getTermType();
        final SarfTermType rightSideCaption = noRightConjugations ? null : rightSideConjugations.getTermType();
        addCaptionRow(leftSideCaption, rightSideCaption);

        ConjugationTuple leftTuple = noLeftConjugations ? null : leftSideConjugations.getMasculineThirdPerson();
        ConjugationTuple rightTuple = noRightConjugations ? null : rightSideConjugations.getMasculineThirdPerson();
        addConjugationRow(leftTuple, rightTuple);

        leftTuple = noLeftConjugations ? null : leftSideConjugations.getFeminineThirdPerson();
        rightTuple = noRightConjugations ? null : rightSideConjugations.getFeminineThirdPerson();
        addConjugationRow(leftTuple, rightTuple);

        leftTuple = noLeftConjugations ? null : leftSideConjugations.getMasculineSecondPerson();
        rightTuple = noRightConjugations ? null : rightSideConjugations.getMasculineSecondPerson();
        addConjugationRow(leftTuple, rightTuple);

        leftTuple = noLeftConjugations ? null : leftSideConjugations.getFeminineSecondPerson();
        rightTuple = noRightConjugations ? null : rightSideConjugations.getFeminineSecondPerson();
        addConjugationRow(leftTuple, rightTuple);

        leftTuple = noLeftConjugations ? null : leftSideConjugations.getFirstPerson();
        rightTuple = noRightConjugations ? null : rightSideConjugations.getFirstPerson();
        addConjugationRow(leftTuple, rightTuple);
        addSeparatorRow(tableAdapter, NUM_OF_COLUMNS);
    }

    private void addNounPair(NounDetailedConjugationPair conjugationPair) {
        if (conjugationPair == null) {
            return;
        }
        final NounConjugationGroup leftSideConjugations = conjugationPair.getLeftSideConjugations();
        final NounConjugationGroup rightSideConjugations = conjugationPair.getRightSideConjugations();
        final boolean noLeftConjugations = leftSideConjugations == null;
        final boolean noRightConjugations = rightSideConjugations == null;

        final SarfTermType leftSideCaption = noLeftConjugations ? null : leftSideConjugations.getTermType();
        final SarfTermType rightSideCaption = noRightConjugations ? null : rightSideConjugations.getTermType();
        addCaptionRow(leftSideCaption, rightSideCaption);

        ConjugationTuple leftTuple = noLeftConjugations ? null : leftSideConjugations.getNominative();
        ConjugationTuple rightTuple = noRightConjugations ? null : rightSideConjugations.getNominative();
        addConjugationRow(leftTuple, rightTuple);

        leftTuple = noLeftConjugations ? null : leftSideConjugations.getAccusative();
        rightTuple = noRightConjugations ? null : rightSideConjugations.getAccusative();
        addConjugationRow(leftTuple, rightTuple);

        leftTuple = noLeftConjugations ? null : leftSideConjugations.getGenitive();
        rightTuple = noRightConjugations ? null : rightSideConjugations.getGenitive();
        addConjugationRow(leftTuple, rightTuple);
        addSeparatorRow(tableAdapter, NUM_OF_COLUMNS);
    }

    private void addNounPairs(NounDetailedConjugationPair[] conjugationPairs) {
        if (isEmpty(conjugationPairs)) {
            return;
        }
        for (NounDetailedConjugationPair conjugationPair : conjugationPairs) {
            addNounPair(conjugationPair);
        }
    }

    private void addCaptionRow(SarfTermType leftSideCaption, SarfTermType rightSideCaption) {
        TcPr leftTcPr = getColumnProperties(leftSideCaption);
        TcPr rightTcPr = getColumnProperties(rightSideCaption);
        final String leftSideCaptionValue = (leftSideCaption == null) ? null : leftSideCaption.toLabel().toUnicode();
        final String rightSideCaptionValue = (rightSideCaption == null) ? null : rightSideCaption.toLabel().toUnicode();
        tableAdapter.startRow()
                .addColumn(0, 3, leftTcPr, getArabicTextPWithStyle(leftSideCaptionValue, ARABIC_CAPTION_STYLE))
                .addColumn(3, (Integer) null, VerticalMergeType.RESTART, getColumnProperties(null), createNoSpacingStyleP())
                .addColumn(4, 3, rightTcPr, getArabicTextPWithStyle(rightSideCaptionValue, ARABIC_CAPTION_STYLE)).endRow();
    }

    private void addConjugationRow(ConjugationTuple leftConjugationTuple, ConjugationTuple rightConjugationTuple) {
        if (leftConjugationTuple == null && rightConjugationTuple == null) {
            return;
        }
        tableAdapter.startRow();

        boolean empty = leftConjugationTuple == null;
        TcPr tcPr = empty ? getNilBorderColumnProperties() : null;
        RootWord rootWord = empty ? null : leftConjugationTuple.getPlural();
        String value = (rootWord == null) ? null : rootWord.toLabel().toUnicode();
        tableAdapter.addColumn(0, tcPr, getArabicTextP(value));

        rootWord = empty ? null : leftConjugationTuple.getDual();
        value = (rootWord == null) ? null : rootWord.toLabel().toUnicode();
        tableAdapter.addColumn(1, tcPr, getArabicTextP(value));

        rootWord = empty ? null : leftConjugationTuple.getSingular();
        value = (rootWord == null) ? null : rootWord.toLabel().toUnicode();
        tableAdapter.addColumn(2, tcPr, getArabicTextP(value));

        tcPr = getNilBorderColumnProperties();
        tableAdapter.addColumn(3, (Integer) null, VerticalMergeType.CONTINUE, tcPr, createNoSpacingStyleP());

        empty = rightConjugationTuple == null;
        tcPr = empty ? getNilBorderColumnProperties() : null;
        rootWord = empty ? null : rightConjugationTuple.getPlural();
        value = (rootWord == null) ? null : rootWord.toLabel().toUnicode();
        tableAdapter.addColumn(4, tcPr, getArabicTextP(value));

        rootWord = empty ? null : rightConjugationTuple.getDual();
        value = (rootWord == null) ? null : rootWord.toLabel().toUnicode();
        tableAdapter.addColumn(5, tcPr, getArabicTextP(value));

        rootWord = empty ? null : rightConjugationTuple.getSingular();
        value = (rootWord == null) ? null : rootWord.toLabel().toUnicode();
        tableAdapter.addColumn(6, tcPr, getArabicTextP(value));

        tableAdapter.endRow();
    }

}
