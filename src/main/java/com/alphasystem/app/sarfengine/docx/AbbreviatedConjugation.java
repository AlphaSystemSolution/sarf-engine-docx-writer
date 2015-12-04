/**
 *
 */
package com.alphasystem.app.sarfengine.docx;

import com.alphasystem.app.sarfengine.conjugation.model.ConjugationHeader;
import com.alphasystem.app.sarfengine.conjugation.model.SarfSagheer;
import com.alphasystem.app.sarfengine.conjugation.model.sarfsagheer.ActiveLine;
import com.alphasystem.app.sarfengine.conjugation.model.sarfsagheer.AdverbLine;
import com.alphasystem.app.sarfengine.conjugation.model.sarfsagheer.ImperativeAndForbiddingLine;
import com.alphasystem.app.sarfengine.conjugation.model.sarfsagheer.PassiveLine;
import com.alphasystem.arabic.model.ArabicLetterType;
import com.alphasystem.arabic.model.ArabicSupport;
import com.alphasystem.arabic.model.ArabicWord;
import com.alphasystem.openxml.builder.TableAdapter;
import com.alphasystem.sarfengine.xml.model.ChartConfiguration;
import org.docx4j.wml.*;

import javax.xml.bind.JAXBElement;

import static com.alphasystem.app.sarfengine.docx.ConjugationHelper.*;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.createCTMarkupRange;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.getText;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;

/**
 * Populate the word document with abbreviated conjugation (Sarf Sagheer).
 *
 * @author sali
 */
public class AbbreviatedConjugation {

    private final SarfSagheer sarfSagheer;

    private final ConjugationHeader conjugationHeader;

    private ChartConfiguration configuration;

    private final TableAdapter tableAdapter;

    /**
     * @param configuration
     * @param sarfSagheer
     * @param conjugationHeader
     */
    public AbbreviatedConjugation(ChartConfiguration configuration,
                                  SarfSagheer sarfSagheer, ConjugationHeader conjugationHeader) {
        this.configuration = configuration;
        this.sarfSagheer = sarfSagheer;
        this.conjugationHeader = conjugationHeader;
        tableAdapter = new TableAdapter(4).startTable();
    }

    private void addActiveLineRow() {
        ActiveLine activeLine = sarfSagheer.getActiveLine();
        if (activeLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, null, null, getArabicTextP(null, activeLine.getActiveParticipleMasculine()))
                    .addColumn(1, null, null, getArabicTextP(null, getMultiWord(activeLine.getVerbalNouns())))
                    .addColumn(2, null, null, getArabicTextP(null, activeLine.getPresentTense()))
                    .addColumn(3, null, null, getArabicTextP(null, activeLine.getPastTense()))
                    .endRow();
        }
    }

    private void addCommandLine() {
        ImperativeAndForbiddingLine commandLine = sarfSagheer.getImperativeAndForbiddingLine();
        if (commandLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, 2, null, getArabicTextP(FORBIDDING_PREFIX, commandLine.getForbidding()))
                    .addColumn(2, 2, null, getArabicTextP(COMMAND_PREFIX, commandLine.getImperative())).endRow();
        }
    }

    private void addHeaderRow() {
        String rsidR = nextId();
        String rsidP = nextId();

        // Root Word
        P rootWordsPara = getRootWordsPara(rsidR, rsidP);

        // translation
        P translationPara = getTranslationPara(rsidR, rsidP);

        // second column paras
        String rsidRpr = nextId();
        P labelP1 = getHeaderLabelPara(rsidR, rsidRpr, rsidP,
                conjugationHeader.getTypeLabel1());
        P labelP2 = getHeaderLabelPara(rsidR, rsidRpr, rsidP,
                conjugationHeader.getTypeLabel2());
        P labelP3 = getHeaderLabelPara(rsidR, rsidRpr, rsidP,
                conjugationHeader.getTypeLabel3());

        tableAdapter.startRow()
                .addColumn(0, 2, null, rootWordsPara, translationPara)
                .addColumn(2, 2, null, labelP1, labelP2, labelP3).endRow();
    }

    private void addPassiveLine() {
        PassiveLine passiveLine = sarfSagheer.getPassiveLine();
        if (passiveLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, null, null, getArabicTextP(null, passiveLine.getPassiveParticipleMasculine()))
                    .addColumn(1, null, null, getArabicTextP(null, getMultiWord(passiveLine.getVerbalNouns())))
                    .addColumn(2, null, null, getArabicTextP(null, passiveLine.getPresentPassiveTense()))
                    .addColumn(3, null, null, getArabicTextP(null, passiveLine.getPastPassiveTense()))
                    .endRow();
        }
    }

    private void addTitleRow() {
        tableAdapter.startRow().addColumn(0, 4, NIL_BORDERS, createTitlePara())
                .endRow();
    }

    private void addZarfLine() {
        AdverbLine zarfLine = sarfSagheer.getAdverbLine();
        if (zarfLine != null) {
            tableAdapter
                    .startRow()
                    .addColumn(0, 4, null, getArabicTextP(ZARF_PREFIX, getMultiWord(zarfLine.getAdverbs())))
                    .endRow();
        }
    }

    /**
     * create title para.
     *
     * @return P
     */
    private P createTitlePara() {
        String id = nextId();

        String bookmarkId = getBookMarkId();
        String bookmarkName = ConjugationHelper.getBookMarkName(bookmarkId);
        CTBookmark bookmarkStart = getCTBookmarkBuilder().withId(bookmarkId)
                .withName(bookmarkName).getObject();
        JAXBElement<CTMarkupRange> bookmarkEnd = createCTMarkupRange(getCTBookmarkRangeBuilder()
                .withId(bookmarkId).getObject());

        RPr rpr = getRPrBuilder().withRFonts(RFONTS_CS)
                .withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        Text text = getText(getTitleWord(sarfSagheer.getActiveLine())
                .toUnicode(), null);
        R r = getRBuilder().withRsidRPr(id).withRPr(rpr).addContent(text)
                .getObject();

        ParaRPr prpr = getParaRPrBuilder().getObject();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_HEADING_STYLE)
                .withBidi(BOOLEAN_DEFAULT_TRUE_TRUE).withRPr(prpr).getObject();

        return getPBuilder().withParaId(id).withRsidP(id).withRsidR(id).withRsidRDefault(id).withRsidRPr(id)
                .withPPr(ppr).addContent(bookmarkStart, r, bookmarkEnd).getObject();

    }

    private P getArabicTextP(ArabicWord prefix, ArabicSupport value) {
        String rsidr = nextId();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_TABLE_CENTER_STYLE)
                .getObject();
        RPr rpr = getRPrBuilder().withRFonts(RFONTS_CS)
                .withRtl(BOOLEAN_DEFAULT_TRUE_TRUE).getObject();
        ArabicWord word = value.getLabel();
        if (prefix != null) {
            word = ArabicWord.concatenateWithSpace(prefix, word);
        }
        Text text = getText(word.toUnicode(), null);
        String id = nextId();
        R r = getRBuilder().withRsidRPr(id).withRPr(rpr).addContent(text)
                .getObject();
        return getPBuilder().withRsidR(rsidr).withRsidRDefault(rsidr)
                .withRsidRPr(id).withRsidP(id).withPPr(ppr).addContent(r)
                .getObject();
    }

    public Tbl getChart() {
        if (!configuration.isOmitTitle()) {
            addTitleRow();
        }
        if (!configuration.isOmitHeader()) {
            addHeaderRow();
        }
        addActiveLineRow();
        addPassiveLine();
        addCommandLine();
        addZarfLine();
        addSeparatorRow(tableAdapter, 4);
        return tableAdapter.getTable();
    }

    private P getHeaderLabelPara(String rsidR, String rsidRpr, String rsidP,
                                 ArabicWord label) {
        ParaRPr prpr = getParaRPrBuilder().withSz(SIZE_32).withSzCs(SIZE_32)
                .getObject();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_NORMAL_STYLE)
                .withBidi(BOOLEAN_DEFAULT_TRUE_TRUE).withRPr(prpr).getObject();

        Text text = getText(label.toUnicode(), null);
        RPr rpr = getRPrBuilder().withRFonts(RFONTS_CS).withSz(SIZE_32)
                .withSzCs(SIZE_32).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text)
                .getObject();

        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR)
                .withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }

    private String getRootWords() {
        StringBuilder rootWords = new StringBuilder();
        if (conjugationHeader != null) {
            ArabicLetterType[] rootLetters = conjugationHeader.getRootLetters();
            rootWords.append(rootLetters[0].toUnicode());
            for (int i = 1; i < rootLetters.length; i++) {
                ArabicLetterType rootLetter = rootLetters[i];
                if (rootLetter == null) {
                    continue;
                }
                rootWords.append(" ").append(rootLetter.toUnicode());
            }
        }
        return rootWords.toString();
    }

    private P getRootWordsPara(String rsidR, String rsidP) {
        ParaRPr prpr = getParaRPrBuilder().withSz(SIZE_56).withSzCs(SIZE_56)
                .getObject();
        PPr ppr = getPPrBuilder().withPStyle(ARABIC_NORMAL_STYLE)
                .withBidi(BOOLEAN_DEFAULT_TRUE_TRUE).withJc(JC_CENTER)
                .withRPr(prpr).getObject();

        Text text = getText(getRootWords(), null);
        RPr rpr = getRPrBuilder().withRFonts(RFONTS_CS).withSz(SIZE_56)
                .withSzCs(SIZE_56).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text)
                .getObject();

        String rsidRpr = nextId();
        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR)
                .withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }

    private P getTranslationPara(String rsidR, String rsidP) {
        String translation = conjugationHeader.getTranslation();
        translation = translation == null ? "" : format("(%s)", translation);
        Text text = getText(translation, null);
        RPr rpr = getRPrBuilder().withRFonts(GEORGIA_FONTS).getObject();
        R r = getRBuilder().withRsidR(rsidR).withRPr(rpr).addContent(text)
                .getObject();
        String rsidRpr = nextId();
        ParaRPr prpr = getParaRPrBuilder().withRFonts(GEORGIA_FONTS)
                .getObject();
        PPr ppr = getPPrBuilder().withJc(JC_CENTER).withRPr(prpr).getObject();
        return getPBuilder().withRsidR(rsidR).withRsidRDefault(rsidR)
                .withRsidP(rsidP).withRsidRPr(rsidRpr).withPPr(ppr)
                .addContent(r).getObject();
    }
}
