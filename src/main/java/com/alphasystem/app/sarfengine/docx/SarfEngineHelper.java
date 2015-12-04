/**
 *
 */
package com.alphasystem.app.sarfengine.docx;

import com.alphasystem.app.sarfengine.conjugation.builder.ConjugationBuilderFactory;
import com.alphasystem.app.sarfengine.conjugation.model.SarfChart;
import com.alphasystem.app.sarfengine.conjugation.model.SarfChartComparator;
import com.alphasystem.app.sarfengine.guice.GuiceSupport;
import com.alphasystem.arabic.model.ArabicLetterType;
import com.alphasystem.arabic.model.NamedTemplate;
import com.alphasystem.sarfengine.xml.model.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.util.List;
import java.util.SortedSet;
import java.util.TreeSet;

import static com.alphasystem.sarfengine.xml.model.SortDirective.NONE;
import static java.util.Collections.synchronizedSortedSet;

/**
 * @author sali
 */
public final class SarfEngineHelper {

    @SuppressWarnings("unused")
    private static final Logger LOGGER = LoggerFactory.getLogger(SarfEngineHelper.class);

    private static final ConjugationBuilderFactory BUILDER_FACTORY = GuiceSupport.getInstance().getConjugationBuilderFactory();

    /**
     * @param template
     * @param removePassiveLine
     * @param skipRuleProcessing
     * @param translation
     * @param firstRadical
     * @param secondRadical
     * @param thirdRadical
     * @param fourthRadical
     * @param verbalNouns
     * @param adverbs
     * @return
     */
    private static SarfChart createSarfChart(NamedTemplate template, boolean removePassiveLine,
                                             boolean skipRuleProcessing, String translation,
                                             ArabicLetterType firstRadical, ArabicLetterType secondRadical,
                                             ArabicLetterType thirdRadical, ArabicLetterType fourthRadical,
                                             List<VerbalNoun> verbalNouns, List<NounOfPlaceAndTime> adverbs) {
        return BUILDER_FACTORY.getConjugationBuilder().doConjugation(template, translation, removePassiveLine,
                skipRuleProcessing, firstRadical, secondRadical, thirdRadical, fourthRadical, verbalNouns, adverbs);
    }

    /**
     * @param template
     * @param removePassiveLine
     * @param skipRuleProcessing
     * @param translation
     * @param firstRadical
     * @param secondRadical
     * @param thirdRadical
     * @param verbalNouns
     * @param adverbs
     * @return
     */
    private static SarfChart createSarfChart(NamedTemplate template,
                                             boolean removePassiveLine, boolean skipRuleProcessing,
                                             String translation, ArabicLetterType firstRadical,
                                             ArabicLetterType secondRadical, ArabicLetterType thirdRadical,
                                             List<VerbalNoun> verbalNouns, List<NounOfPlaceAndTime> adverbs) {
        return createSarfChart(template, removePassiveLine, skipRuleProcessing, translation, firstRadical, secondRadical,
                thirdRadical, null, verbalNouns, adverbs);
    }

    /**
     * @param file
     * @param configuration
     * @param charts
     */
    public synchronized static void execute(File file,
                                            ChartConfiguration configuration, SortedSet<SarfChart> charts) {
        SarfChartComparator chartComparator = new SarfChartComparator(
                configuration.getSortDirective(),
                configuration.getSortDirection());
        SortedSet<SarfChart> sc = synchronizedSortedSet(new TreeSet<>(chartComparator));
        sc.addAll(charts);
        SarfEngine sarfEngine = new SarfEngine(file, configuration, sc.toArray(new SarfChart[sc.size()]));
        sarfEngine.execute();

    }

    /**
     * @param file
     * @param charts
     */
    public synchronized static void execute(File file, SortedSet<SarfChart> charts) {
        execute(file, null, charts);
    }

    private SortedSet<SarfChart> charts;

    public SarfEngineHelper() {
        this(null);
    }

    public SarfEngineHelper(SarfChartComparator comparator) {
        init(comparator);
    }

    /**
     * @param template
     * @param removePassiveLine
     * @param skipRuleProcessing
     * @param translation
     * @param firstRadical
     * @param secondRadical
     * @param thirdRadical
     * @param fourthRadical
     * @param verbalNouns
     * @param adverbs
     */
    public void add(NamedTemplate template, boolean removePassiveLine, boolean skipRuleProcessing, String translation,
                    ArabicLetterType firstRadical, ArabicLetterType secondRadical, ArabicLetterType thirdRadical,
                    ArabicLetterType fourthRadical, List<VerbalNoun> verbalNouns, List<NounOfPlaceAndTime> adverbs) {
        charts.add(createSarfChart(template, removePassiveLine, skipRuleProcessing, translation, firstRadical,
                secondRadical, thirdRadical, fourthRadical, verbalNouns, adverbs));
    }

    /**
     * @param template
     * @param removePassiveLine
     * @param skipRuleProcessing
     * @param translation
     * @param firstRadical
     * @param secondRadical
     * @param thirdRadical
     * @param verbalNouns
     * @param adverbs
     */
    public void add(NamedTemplate template, boolean removePassiveLine, boolean skipRuleProcessing, String translation,
                    ArabicLetterType firstRadical, ArabicLetterType secondRadical, ArabicLetterType thirdRadical,
                    List<VerbalNoun> verbalNouns, List<NounOfPlaceAndTime> adverbs) {
        add(template, removePassiveLine, skipRuleProcessing, translation, firstRadical, secondRadical, thirdRadical,
                null, verbalNouns, adverbs);
    }

    /**
     * @param template
     * @param firstRadical
     * @param secondRadical
     * @param thirdRadical
     * @param verbalNouns
     * @param adverbs
     */
    public void add(NamedTemplate template, ArabicLetterType firstRadical, ArabicLetterType secondRadical,
                    ArabicLetterType thirdRadical, List<VerbalNoun> verbalNouns, List<NounOfPlaceAndTime> adverbs) {
        add(template, false, false, null, firstRadical, secondRadical, thirdRadical, verbalNouns, adverbs);
    }

    /**
     * @param template
     */
    public void addAll(ConjugationTemplate template) {
        List<ConjugationData> data = template.getData();
        for (ConjugationData cd : data) {
            ConjugationConfiguration configuration = cd.getConfiguration();
            RootLetters rootLetters = cd.getRootLetters();
            add(cd.getTemplate(), configuration.isRemovePassiveLine(), configuration.isSkipRuleProcessing(),
                    cd.getTranslation(), rootLetters.getFirstRadical(), rootLetters.getSecondRadical(),
                    rootLetters.getThirdRadical(), rootLetters.getFourthRadical(), cd.getVerbalNouns(), cd.getAdverbs());
        }
    }

    /**
     * @param file
     */
    public synchronized void execute(File file) {
        execute(file, new ChartConfiguration(), charts);
    }

    /**
     * @param file
     * @param configuration
     */
    public synchronized void execute(File file, ChartConfiguration configuration) {
        execute(file, configuration, charts);
    }

    /**
     * @return
     */
    public SortedSet<SarfChart> getCharts() {
        return charts;
    }

    /**
     * @param comparator
     */
    private void init(SarfChartComparator comparator) {
        SarfChartComparator c = comparator == null ? new SarfChartComparator(NONE) : comparator;
        charts = synchronizedSortedSet(new TreeSet<>(c));
    }
}
