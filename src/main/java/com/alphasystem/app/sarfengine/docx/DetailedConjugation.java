/**
 *
 */
package com.alphasystem.app.sarfengine.docx;

import com.alphasystem.app.sarfengine.conjugation.model.SarfKabeer;
import com.alphasystem.app.sarfengine.conjugation.model.SarfTerm;
import com.alphasystem.arabic.model.ArabicWord;
import com.alphasystem.openxml.builder.TableAdapter;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.TcPrInner.TcBorders;

import java.util.List;

import static com.alphasystem.app.sarfengine.docx.ConjugationHelper.*;
import static com.alphasystem.arabic.model.ArabicLetters.WORD_SPACE;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.getText;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static org.apache.commons.lang3.ArrayUtils.subarray;

/**
 * @author sali
 */
public class DetailedConjugation {

    private final List<SarfTerm> sarfTerms;

    private final TableAdapter tableAdapter;

    /**
     * @param sarfTerms
     */
    public DetailedConjugation(List<SarfTerm> sarfTerms) {
        this.sarfTerms = sarfTerms;
        tableAdapter = new TableAdapter(7)
                .setColumnsWidth(new int[]{0, 1, 2, 4, 5, 6}, 16.24)
                .setColumnWidth(3, 2.56).startTable();
    }

    public DetailedConjugation(SarfKabeer sarfKabeer) {
        this(fromSarfKabeer(sarfKabeer));
    }

    private void addCaptionRow(ArabicWord rightSideCaption,
                               ArabicWord leftSideCaption, boolean noBorder) {
        tableAdapter
                .startRow()
                .addColumn(0, 3, noBorder ? NIL_BORDERS : null,
                        getArabicTermP(rightSideCaption, ARABIC_CAPTION_STYLE))
                .addColumn(3, null, NIL_BORDERS, createNoSpacingStyleP())
                .addColumn(4, 3, null,
                        getArabicTermP(leftSideCaption, ARABIC_CAPTION_STYLE))
                .endRow();
    }

    private void addConjugationRow(ArabicWord[] rightSideValues,
                                   ArabicWord[] leftSideValues, boolean noBorder) {
        tableAdapter.startRow();
        int columnIndex = 0;
        for (int i = 0; i < rightSideValues.length; i++) {
            TcBorders borders = noBorder ? NIL_BORDERS : null;
            P p = getArabicTermP(rightSideValues[i], ARABIC_TABLE_CENTER_STYLE);
            tableAdapter.addColumn(columnIndex, 1, borders, p);
            columnIndex++;
        }
        tableAdapter.addColumn(columnIndex, 1, NIL_BORDERS,
                createNoSpacingStyleP());
        columnIndex++;
        for (int i = 0; i < leftSideValues.length; i++) {
            P p = getArabicTermP(leftSideValues[i], ARABIC_TABLE_CENTER_STYLE);
            tableAdapter.addColumn(columnIndex, 1, null, p);
            columnIndex++;
        }
        tableAdapter.endRow();
    }

    /**
     * @param rightSideTerm
     * @param leftSideTerm
     */
    private void addConjugationRows(SarfTerm rightSideTerm,
                                    SarfTerm leftSideTerm) {
        boolean noBorder = rightSideTerm == null
                || rightSideTerm.getSarfTermType() == null;
        ArabicWord rightSideCaption = noBorder ? null : rightSideTerm
                .getLabel();
        ArabicWord leftSideCaption = leftSideTerm.getLabel();
        addCaptionRow(rightSideCaption, leftSideCaption, noBorder);

        ArabicWord[] rightSideValues = noBorder ? null : rightSideTerm
                .getValues();
        ArabicWord[] leftSideValues = leftSideTerm.getValues();

        int fromIndex = 0;
        int toIndex = 3;
        while (fromIndex < leftSideValues.length) {
            ArabicWord[] rightSideSubValues = noBorder ? new ArabicWord[3]
                    : subarray(rightSideValues, fromIndex, toIndex);
            ArabicWord[] leftSideSubValues = subarray(leftSideValues,
                    fromIndex, toIndex);
            addConjugationRow(rightSideSubValues, leftSideSubValues, noBorder);
            fromIndex = toIndex;
            toIndex += 3;
        }

        addSeparatorRow(tableAdapter, 7);
    }

    private P getArabicTermP(ArabicWord arabicWord, PStyle style) {
        PPr ppr = getPPrBuilder().withPStyle(style).getObject();

        RPr rpr = getRPrBuilder().withRFonts(RFONTS_CS)
                .withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        String value = arabicWord == null ? WORD_SPACE.toUnicode() : arabicWord
                .toUnicode();
        Text text = getText(value, null);
        String rsid = nextId();
        R r = getRBuilder().withRsidRPr(rsid).withRPr(rpr).addContent(text)
                .getObject();
        String id = nextId();
        return getPBuilder().withRsidR(id).withRsidRDefault(id).withRsidP(rsid)
                .withRsidRPr(rsid).withPPr(ppr).addContent(r).getObject();
    }

    /**
     * @return table containing detail conjugation (Sarf Kabeer)
     */
    public Tbl getChart() {
        int fromIndex = 0;
        int toIndex = 2;
        while (fromIndex < sarfTerms.size()) {
            List<SarfTerm> subList = sarfTerms.subList(fromIndex, toIndex);
            addConjugationRows(subList.get(1), subList.get(0));
            fromIndex = toIndex;
            toIndex += 2;
        }

        addSeparatorRow(tableAdapter, 7);

        return tableAdapter.getTable();
    }

}
