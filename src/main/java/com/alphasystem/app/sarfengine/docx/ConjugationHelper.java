/**
 *
 */
package com.alphasystem.app.sarfengine.docx;

import com.alphasystem.app.sarfengine.conjugation.model.ConjugationStack;
import com.alphasystem.app.sarfengine.conjugation.model.SarfKabeer;
import com.alphasystem.app.sarfengine.conjugation.model.SarfKabeerPair;
import com.alphasystem.app.sarfengine.conjugation.model.SarfTerm;
import com.alphasystem.app.sarfengine.conjugation.model.sarfsagheer.ActiveLine;
import com.alphasystem.arabic.model.ArabicSupport;
import com.alphasystem.arabic.model.ArabicWord;
import com.alphasystem.openxml.builder.TableAdapter;
import com.alphasystem.sarfengine.xml.model.RootWord;
import com.alphasystem.sarfengine.xml.model.SarfTermType;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.PStyle;
import org.docx4j.wml.TcPrInner.TcBorders;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

import static com.alphasystem.arabic.model.ArabicLetterType.*;
import static com.alphasystem.arabic.model.ArabicLetters.WORD_SPACE;
import static com.alphasystem.arabic.model.ArabicWord.*;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.getHpsMeasure;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.getPStyle;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.*;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.lang.String.format;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.docx4j.wml.STBorder.NIL;
import static org.docx4j.wml.STHint.CS;

/**
 * @author sali
 */
public class ConjugationHelper {

    public static final RFonts RFONTS_CS = getRFontsBuilder().withHint(CS).getObject();

    public static final HpsMeasure SIZE_56 = getHpsMeasure("56");

    public static final HpsMeasure SIZE_32 = getHpsMeasure("32");

    public static final CTBorder NIL_BORDER = getCTBorderBuilder().withVal(NIL).getObject();

    public static final TcBorders NIL_BORDERS = getTcPrInnerTcBordersBuilder()
            .withTop(NIL_BORDER).withBottom(NIL_BORDER).withLeft(NIL_BORDER)
            .withRight(NIL_BORDER).getObject();

    public static final RFonts GEORGIA_FONTS = getRFontsBuilder()
            .withAscii("Georgia").withHAnsi("Georgia").getObject();

    public static final PStyle NO_SPACING_STYLE = getPStyle("NoSpacing");

    public static final PStyle ARABIC_NORMAL_STYLE = getPStyle("Arabic-Normal");

    public static final PStyle ARABIC_TABLE_CENTER_STYLE = getPStyle("Arabic-Table-Center");

    public static final PStyle ARABIC_CAPTION_STYLE = getPStyle("Arabic-Caption");

    public static final PStyle ARABIC_HEADING_STYLE = getPStyle("Arabic-Heading1");

    public static final PStyle ARABIC_TOC_STYLE = getPStyle("TOC1");

    public static final ArabicWord COMMAND_PREFIX = getWord(ALIF, LAM,
            ALIF_HAMZA_ABOVE, MEEM, RA, SPACE, MEEM, NOON, HA);

    public static final ArabicWord FORBIDDING_PREFIX = getWord(WAW, NOON, HA,
            YA, SPACE, AIN, NOON, HA);

    public static final ArabicWord ZARF_PREFIX = getWord(WAW, ALIF, LAM, DTHA,
            RA, FA, SPACE, MEEM, NOON, HA);

    public static final String BOOKMARK_NAME_PREFIX = "bm";

    public static final AtomicInteger BOOKMARK_COUNT = new AtomicInteger(0);

    /**
     * Add empty separator row.
     */
    public static void addSeparatorRow(TableAdapter tableAdapter,
                                       Integer gridSpan) {
        tableAdapter.startRow()
                .addColumn(0, gridSpan, NIL_BORDERS, createNoSpacingStyleP())
                .endRow();
    }

    public static P createNoSpacingStyleP() {
        PPr ppr = getPPrBuilder().withPStyle(NO_SPACING_STYLE).getObject();
        P p = getPBuilder().withRsidR(nextId()).withRsidP(nextId())
                .withRsidRDefault(nextId()).withPPr(ppr).getObject();
        return p;
    }

    public static String getBookMarkId() {
        return format("%s", BOOKMARK_COUNT.incrementAndGet());
    }

    public static String getBookMarkName(String id) {
        return format("%s_%s", BOOKMARK_NAME_PREFIX, id);
    }

    public static List<SarfTerm> fromSarfKabeer(SarfKabeer sarfKabeer) {
        List<SarfTerm> sarfTerms = new ArrayList<>();
        if (sarfKabeer != null) {
            loadTerms(sarfTerms, sarfKabeer.getActiveTensePair());
            loadTerms(sarfTerms, sarfKabeer.getVerbalNounPairs());
            loadTerms(sarfTerms, sarfKabeer.getActiveParticiplePair());
            loadTerms(sarfTerms, sarfKabeer.getPassiveTensePair());
            loadTerms(sarfTerms, sarfKabeer.getPassiveParticiplePair());
            loadTerms(sarfTerms, sarfKabeer.getImperativeAndForbiddingPair());
            loadTerms(sarfTerms, sarfKabeer.getAdverbPairs());
        }
        return sarfTerms;
    }

    public static ArabicWord getMultiWord(ArabicWord[] words) {
        ArabicWord w = WORD_SPACE;
        if (!isEmpty(words)) {
            w = words[0];
            for (int i = 1; i < words.length; i++) {
                w = concatenateWithAnd(w, words[i]);
            }
        }
        return w;
    }

    /**
     * Gets the title of the conjugation. The title will be comprised of third
     * person singular masculine past tense (space> third person singular
     * masculine present tense.
     *
     * @param activeLine
     * @return
     */
    public static ArabicWord getTitleWord(ActiveLine activeLine) {
        ArabicWord pastTense = WORD_SPACE;
        ArabicWord presentTense = WORD_SPACE;
        if (activeLine != null) {
            RootWord pastTenseRootWord = activeLine.getPastTense();
            pastTense = pastTenseRootWord == null ? WORD_SPACE : pastTenseRootWord.getRootWord();
            if (pastTense == null) {
                pastTense = WORD_SPACE;
            }
            RootWord presentTenseRootWord = activeLine.getPresentTense();
            presentTense = presentTenseRootWord == null ? WORD_SPACE : presentTenseRootWord.getRootWord();
            if (presentTense == null) {
                presentTense = WORD_SPACE;
            }
        }
        return concatenateWithSpace(pastTense, presentTense);
    }

    private static void loadTerms(List<SarfTerm> sarfTerms, SarfKabeerPair pair) {
        if (pair != null) {
            ConjugationStack leftSideStack = pair.getLeftSideStack();
            ConjugationStack rightSideStack = pair.getRightSideStack();
            RootWord[] rightSideConjugations = null;
            SarfTermType rightSideLabel = null;
            RootWord rightSideDefaultValue = null;
            if (rightSideStack != null) {
                rightSideConjugations = rightSideStack.getConjugations();
                rightSideLabel = rightSideStack.getLabel();
                rightSideDefaultValue = rightSideStack.getDefaultValue();
            }
            int size = isEmpty(rightSideConjugations) ? 0 : rightSideConjugations.length;

            SarfTerm rightSideTerm = createSarfTerm(rightSideLabel, rightSideDefaultValue, rightSideConjugations, size);
            SarfTerm leftSideTerm = createSarfTerm(leftSideStack.getLabel(), leftSideStack.getDefaultValue(),
                    leftSideStack.getConjugations(), size);
            sarfTerms.add(rightSideTerm);
            sarfTerms.add(leftSideTerm);
        }
    }

    private static void loadTerms(List<SarfTerm> sarfTerms, SarfKabeerPair[] pairs) {
        if (pairs == null) {
            return;
        }
        for (SarfKabeerPair sarfKabeerPair : pairs) {
            loadTerms(sarfTerms, sarfKabeerPair);
        }
    }

    private static SarfTerm createSarfTerm(SarfTermType sarfTermType, RootWord defaultValue, RootWord[] conjugations,
                                           int size) {
        if (size <= 0) {
            // nothing to create return null;
            return null;
        }

        ArabicWord label = defaultValue == null ? null : defaultValue.getLabel();
        return new SarfTerm(sarfTermType, label, getFromArabicSupport(conjugations));
    }

    private static ArabicWord[] getFromArabicSupport(ArabicSupport[] srcValues) {
        if (srcValues == null) {
            return null;
        }
        ArabicWord[] words = new ArabicWord[srcValues.length];
        for (int i = 0; i < words.length; i++) {
            ArabicSupport srcValue = srcValues[i];
            words[i] = srcValue == null ? null : srcValue.getLabel();
        }
        return words;
    }

    public static ArabicWord getMultiWord(RootWord[] rootWords) {
        int length = isEmpty(rootWords) ? 0 : rootWords.length;
        ArabicWord[] words = new ArabicWord[length];
        if (length > 0) {
            for (int i = 0; i < length; i++) {
                RootWord rootWord = rootWords[i];
                words[i] = (rootWord == null) ? null : rootWord.getRootWord();
            }
        }
        ArabicWord w = WORD_SPACE;
        if (!isEmpty(words)) {
            w = words[0];
            for (int i = 1; i < words.length; i++) {
                ArabicWord word = words[i];
                if (word == null) {
                    continue;
                }
                w = concatenateWithAnd(w, word);
            }
        }
        return w;
    }
}
