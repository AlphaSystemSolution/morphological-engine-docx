package com.alphasystem.app.morphologicalengine.docx;

import com.alphasystem.morphologicalanalysis.morphology.model.support.SarfTermType;
import com.alphasystem.morphologicalengine.model.ConjugationTuple;
import com.alphasystem.morphologicalengine.model.DetailedConjugation;
import com.alphasystem.morphologicalengine.model.NounConjugationGroup;
import com.alphasystem.morphologicalengine.model.VerbConjugationGroup;
import com.alphasystem.openxml.builder.wml.table.TableAdapter;
import com.alphasystem.openxml.builder.wml.table.TableAdapter.VerticalMergeType;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.TcPr;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import static com.alphasystem.app.morphologicalengine.docx.WmlHelper.*;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;

/**
 * @author sali
 */
public final class DetailedConjugationAdapter extends ChartAdapter {

    private static final int NUM_OF_COLUMNS = 7;

    private final DetailedConjugation[] detailedConjugations;
    private final TableAdapter tableAdapter;

    DetailedConjugationAdapter(DetailedConjugation... detailedConjugations) {
        this.detailedConjugations = isEmpty(detailedConjugations) ? new DetailedConjugation[0] : detailedConjugations;
        this.tableAdapter = new TableAdapter().startTable(16.24, 16.24, 16.24, 2.56, 16.24, 16.24, 16.24);
    }

    private static TcPr getColumnProperties(Object obj) {
        return (obj == null) ? getNilBorderColumnProperties() : null;
    }

    @Override
    protected Tbl getChart() {
        for (DetailedConjugation detailedConjugation : detailedConjugations) {
            addTensePair(detailedConjugation.getPresentTense(), detailedConjugation.getPastTense());
            addNounPair(detailedConjugation.getActiveParticipleFeminine(), detailedConjugation.getActiveParticipleMasculine());
            addNounPairs(detailedConjugation.getVerbalNouns());
            addTensePair(detailedConjugation.getPresentPassiveTense(), detailedConjugation.getPastPassiveTense());
            addNounPair(detailedConjugation.getPassiveParticipleFeminine(), detailedConjugation.getPassiveParticipleMasculine());
            addTensePair(detailedConjugation.getForbidding(), detailedConjugation.getImperative());
            addNounPairs(detailedConjugation.getAdverbs());
        }

        return tableAdapter.getTable();
    }

    private void addTensePair(final VerbConjugationGroup leftSideConjugations, final VerbConjugationGroup rightSideConjugations) {
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

    private void addNounPair(final NounConjugationGroup leftSideConjugations, final NounConjugationGroup rightSideConjugations) {
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

    private void addNounPairs(NounConjugationGroup[] conjugationPairs) {
        if (isEmpty(conjugationPairs)) {
            return;
        }
        List<NounConjugationGroup> list = new ArrayList<>();
        Collections.addAll(list, conjugationPairs);
        while (list.size() % 2 != 0) {
            list.add(null);
        }
        int from = 0;
        int to = 2;
        while (from < list.size()) {
            final List<NounConjugationGroup> subList = list.subList(from, to);
            addNounPair(subList.get(1), subList.get(0));
            from = to;
            to += 2;
        }
    }

    private void addCaptionRow(SarfTermType leftSideCaption, SarfTermType rightSideCaption) {
        TcPr leftTcPr = getColumnProperties(leftSideCaption);
        TcPr rightTcPr = getColumnProperties(rightSideCaption);
        final String leftSideCaptionValue = (leftSideCaption == null) ? null : leftSideCaption.toLabel().toUnicode();
        final String rightSideCaptionValue = (rightSideCaption == null) ? null : rightSideCaption.toLabel().toUnicode();
        tableAdapter
                .startRow()
                .addColumn(0, 3, leftTcPr, getArabicTextPWithStyle(leftSideCaptionValue, ARABIC_CAPTION_STYLE))
                .addColumn(3, (Integer) null, VerticalMergeType.RESTART, getColumnProperties(null), createNoSpacingStyleP())
                .addColumn(4, 3, rightTcPr, getArabicTextPWithStyle(rightSideCaptionValue, ARABIC_CAPTION_STYLE))
                .endRow();
    }

    private void addConjugationRow(ConjugationTuple leftConjugationTuple, ConjugationTuple rightConjugationTuple) {
        if (leftConjugationTuple == null && rightConjugationTuple == null) {
            return;
        }

        tableAdapter.startRow();
        int columnIndex = addConjugationColumns(leftConjugationTuple, 0);
        tableAdapter.addColumn(columnIndex += 1, (Integer) null, VerticalMergeType.CONTINUE, getNilBorderColumnProperties(),
                createNoSpacingStyleP());
        addConjugationColumns(rightConjugationTuple, columnIndex + 1);
        tableAdapter.endRow();
    }

    private int addConjugationColumns(ConjugationTuple conjugationTuple, int beginColumnIndex) {
        int columnIndex = beginColumnIndex;
        if (conjugationTuple == null) {
            tableAdapter.addColumn(columnIndex, getNilBorderColumnProperties(), getArabicTextP(null));
            tableAdapter.addColumn(columnIndex += 1, getNilBorderColumnProperties(), getArabicTextP(null));
            tableAdapter.addColumn(columnIndex += 1, getNilBorderColumnProperties(), getArabicTextP(null));
        } else {
            String dualValue = conjugationTuple.getDual();
            int gridSpan = dualValue == null ? 2 : 1;
            tableAdapter.addColumn(columnIndex, gridSpan, getArabicTextP(conjugationTuple.getPlural()));

            if (dualValue != null) {
                tableAdapter.addColumn(columnIndex += 1, getArabicTextP(dualValue));
            }

            tableAdapter.addColumn(columnIndex += gridSpan, getArabicTextP(conjugationTuple.getSingular()));
        }
        return columnIndex;
    }

}
