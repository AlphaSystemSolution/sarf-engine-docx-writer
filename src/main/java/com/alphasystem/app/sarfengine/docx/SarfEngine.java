/**
 * 
 */
package com.alphasystem.app.sarfengine.docx;

import static com.alphasystem.app.sarfengine.docx.ConjugationHelper.ARABIC_TOC_STYLE;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.createNewDoc;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.getHpsMeasure;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.getText;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.getWrappedFldChar;
import static com.alphasystem.openxml.builder.OpenXmlAdapter.save;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.BOOLEAN_DEFAULT_TRUE_TRUE;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getCTColumnsBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getCTDocGridBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getCTLanguageBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getCTTabStopBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getFldCharBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getPBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getPPrBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getParaRPrBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getRBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getRFontsBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getSectPrBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getSectPrPgMarBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getSectPrPgSzBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getSectPrTypeBuilder;
import static com.alphasystem.openxml.builder.OpenXmlBuilderFactory.getTabsBuilder;
import static com.alphasystem.util.IdGenerator.nextId;
import static java.util.concurrent.Executors.newSingleThreadExecutor;
import static org.apache.commons.lang3.ArrayUtils.isEmpty;
import static org.docx4j.wml.STFldCharType.BEGIN;
import static org.docx4j.wml.STFldCharType.END;
import static org.docx4j.wml.STTabJc.RIGHT;
import static org.docx4j.wml.STTabTlc.DOT;
import static org.docx4j.wml.STTheme.MINOR_EAST_ASIA;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Future;

import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.CTColumns;
import org.docx4j.wml.CTDocGrid;
import org.docx4j.wml.CTLanguage;
import org.docx4j.wml.CTTabStop;
import org.docx4j.wml.FldChar;
import org.docx4j.wml.HpsMeasure;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.ParaRPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RFonts;
import org.docx4j.wml.SectPr;
import org.docx4j.wml.SectPr.PgMar;
import org.docx4j.wml.SectPr.PgSz;
import org.docx4j.wml.SectPr.Type;
import org.docx4j.wml.Tabs;
import org.docx4j.wml.Text;

import com.alphasystem.ApplicationException;
import com.alphasystem.app.sarfengine.conjugation.model.SarfChart;
import com.alphasystem.app.sarfengine.conjugation.model.SarfKabeer;
import com.alphasystem.sarfengine.xml.model.ChartConfiguration;

/**
 * @author sali
 * 
 */
public class SarfEngine implements Callable<Boolean> {

	private static void addFieldBegin(P paragraph) {
		FldChar fldchar = getFldCharBuilder().withDirty(true)
				.withFldCharType(BEGIN).getObject();
		R r = getRBuilder().addContent(getWrappedFldChar(fldchar)).getObject();
		paragraph.getContent().add(r);
	}

	private static void addFieldEnd(P paragraph) {
		FldChar fldchar = getFldCharBuilder().withFldCharType(END).getObject();
		R r = getRBuilder().addContent(getWrappedFldChar(fldchar)).getObject();
		paragraph.getContent().add(r);
	}

	private static void addTableOfContentField(P paragraph) {
		Text txt = getText("TOC \\o \"1-3\" \\h \\z \\t \"Arabic-Heading1,1\"",
				"preserve");
		R r = getRBuilder().addContent(txt).getObject();
		paragraph.getContent().add(r);
	}

	private static SectPr createSectPrInternal(String rsidR, String num) {
		Type type = getSectPrTypeBuilder().withVal("continuous").getObject();
		PgSz pgSz = getSectPrPgSzBuilder().withW("12240").withH("15840")
				.getObject();
		String value1 = "1440";
		String value2 = "708";
		PgMar pgMar = getSectPrPgMarBuilder().withTop(value1).withRight(value1)
				.withBottom(value1).withLeft(value1).withHeader(value2)
				.withFooter(value2).withGutter("0").getObject();
		CTColumns cols = getCTColumnsBuilder().withNum(num).withSpace("708")
				.getObject();
		CTDocGrid docGrid = getCTDocGridBuilder().withLinePitch("360")
				.getObject();

		return getSectPrBuilder().withRsidR(rsidR).withRsidSect(nextId())
				.withType(type).withPgSz(pgSz).withPgMar(pgMar).withCols(cols)
				.withDocGrid(docGrid).getObject();
	}

	private static SarfChart[] initFromSarfKabeer(SarfKabeer... sarfKabeers) {
		SarfChart[] sarfCharts = new SarfChart[0];
		if (!isEmpty(sarfKabeers)) {
			List<SarfChart> list = new ArrayList<SarfChart>();
			for (SarfKabeer sk : sarfKabeers) {
				list.add(new SarfChart(null, null, sk));
			}
			sarfCharts = list.toArray(new SarfChart[0]);
		}
		return sarfCharts;
	}

	private final SarfChart[] sarfCharts;

	private File file;

	private ChartConfiguration configuration;

	/**
	 * @param destFile
	 * @param sarfCharts
	 */
	public SarfEngine(File destFile, SarfChart... sarfCharts) {
		this(destFile, null, sarfCharts);
		if (isEmpty(this.sarfCharts) || this.sarfCharts.length <= 1) {
			configuration.setOmitToc(true);
		}
	}

	/**
	 * @param destFile
	 * @param configuration
	 * @param sarfCharts
	 */
	public SarfEngine(File destFile, ChartConfiguration configuration,
			SarfChart... sarfCharts) {
		this.sarfCharts = isEmpty(sarfCharts) ? new SarfChart[0] : sarfCharts;
		this.configuration = configuration == null ? new ChartConfiguration()
				: configuration;
		this.file = destFile;
	}

	/**
	 * @param destFil
	 * @param sarfKabeers
	 */
	public SarfEngine(File destFil, SarfKabeer... sarfKabeers) {
		this(destFil, initFromSarfKabeer(sarfKabeers));
	}

	/**
	 * 
	 * @param sarfCharts
	 */
	public SarfEngine(SarfChart... sarfCharts) {
		this(null, null, sarfCharts);
	}

	/**
	 * 
	 * @param configuration
	 * @param sarfCharts
	 */
	public SarfEngine(ChartConfiguration configuration,
			SarfChart... sarfCharts) {
		this(null, configuration, sarfCharts);
	}

	/**
	 * @param configuration
	 * @param sarfKabeers
	 */
	public SarfEngine(ChartConfiguration configuration,
			SarfKabeer... sarfKabeers) {
		this(null, configuration, initFromSarfKabeer(sarfKabeers));
	}

	/**
	 * 
	 * @param sarfKabeers
	 */
	public SarfEngine(SarfKabeer... sarfKabeers) {
		this(null, null, initFromSarfKabeer(sarfKabeers));
	}

	private void buildSarfChart(MainDocumentPart mainDocumentPart) {
		if (isEmpty(sarfCharts)) {
			return;
		}
		if (!configuration.isOmitAbbreviatedConjugation()
				&& !configuration.isOmitToc()) {
			mainDocumentPart.addObject(createTocSectionBreak());
			mainDocumentPart.addObject(createToc());
			mainDocumentPart.addObject(createSecondSectionBreak());
		}
		for (SarfChart sarfChart : sarfCharts) {
			MainConjugation mainConjugation = new MainConjugation(sarfChart,
					configuration);
			try {
				mainConjugation.convert(mainDocumentPart);
			} catch (ApplicationException e) {
				e.printStackTrace();
			}
		}
	}

	@Override
	public Boolean call() throws Exception {
		boolean done;
		boolean exceptionOccured = false;
		try {
			convert();
		} catch (Exception e) {
			exceptionOccured = true;
			throw e;
		} finally {
			done = !exceptionOccured;
		}
		return done;
	}

	public void convert() throws Docx4JException {
		WordprocessingMLPackage wordprocessingMLPackage = createNewDoc();
		MainDocumentPart mainDocumentPart = wordprocessingMLPackage
				.getMainDocumentPart();
		buildSarfChart(mainDocumentPart);
		save(file, wordprocessingMLPackage);
	}

	private P createSecondSectionBreak() {
		String rsidR = nextId();
		SectPr sectPr = createSectPrInternal(rsidR, "2");
		PPr pPr = getPPrBuilder().withSectPr(sectPr).getObject();
		R r = getRBuilder().addContent(
				getFldCharBuilder().withFldCharType(END).getObject())
				.getObject();
		return getPBuilder().withParaId(nextId()).withRsidR(rsidR)
				.withRsidRDefault(nextId()).withRsidP(nextId()).withPPr(pPr)
				.addContent(r).getObject();
	}

	private P createToc() {
		String rsidR = nextId();

		CTTabStop cTTabStop = getCTTabStopBuilder().withVal(RIGHT)
				.withLeader(DOT).withPos("3600").getObject();

		Tabs tabs = getTabsBuilder().addTab(cTTabStop).getObject();

		RFonts rFonts = getRFontsBuilder().withEastAsiaTheme(MINOR_EAST_ASIA)
				.getObject();
		HpsMeasure hpsMeasure = getHpsMeasure(22);
		CTLanguage cTLanguage = getCTLanguageBuilder().withVal("en-US")
				.getObject();
		ParaRPr rPr = getParaRPrBuilder().withRFonts(rFonts)
				.withNoProof(BOOLEAN_DEFAULT_TRUE_TRUE).withSz(hpsMeasure)
				.withSzCs(hpsMeasure).withLang(cTLanguage).getObject();

		PPr ppr = getPPrBuilder().withPStyle(ARABIC_TOC_STYLE).withRPr(rPr)
				.withTabs(tabs).getObject();

		P p = getPBuilder().withParaId(nextId()).withRsidR(rsidR)
				.withRsidRDefault(nextId()).withRsidP(nextId()).withPPr(ppr)
				.getObject();
		addFieldBegin(p);
		addTableOfContentField(p);
		addFieldEnd(p);
		return p;
	}

	private P createTocSectionBreak() {
		String rsidR = nextId();

		SectPr sectPr = createSectPrInternal(rsidR, null);

		CTTabStop cTTabStop = getCTTabStopBuilder().withVal(RIGHT)
				.withLeader(DOT).withPos("9350").getObject();

		Tabs tabs = getTabsBuilder().addTab(cTTabStop).getObject();
		PPr ppr = getPPrBuilder().withPStyle(ARABIC_TOC_STYLE).withTabs(tabs)
				.withSectPr(sectPr).getObject();

		return getPBuilder().withParaId(nextId()).withRsidR(rsidR)
				.withRsidRDefault(nextId()).withRsidP(nextId()).withPPr(ppr)
				.getObject();
	}

	public Boolean execute() {
		Boolean result = true;
		ExecutorService executorService = newSingleThreadExecutor();
		Future<Boolean> done = executorService.submit(this);
		if (done != null) {
			try {
				result = done.get();
			} catch (Exception e) {
				result = false;
				e.printStackTrace();
			} finally {
				executorService.shutdown();
			}
		}
		return result;
	}

	public ChartConfiguration getConfiguration() {
		return configuration;
	}

	public File getFile() {
		return file;
	}

	public void setFile(File file) {
		this.file = file;
	}

}
