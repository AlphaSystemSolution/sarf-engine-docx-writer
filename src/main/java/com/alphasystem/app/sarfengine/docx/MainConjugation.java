/**
 * 
 */
package com.alphasystem.app.sarfengine.docx;

import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;

import com.alphasystem.ApplicationException;
import com.alphasystem.BusinessException;
import com.alphasystem.app.sarfengine.conjugation.model.SarfChart;
import com.alphasystem.sarfengine.xml.model.ChartConfiguration;

/**
 * Builds the entire Sarf chart.
 * 
 * @author sali
 * 
 */
public class MainConjugation {

	protected final AbbreviatedConjugation abbreviatedConjugation;

	protected final DetailedConjugation detailedConjugation;

	protected final SarfChart sarfChart;

	protected ChartConfiguration configuration;

	/**
	 * 
	 * @param sarfChart
	 */
	public MainConjugation(SarfChart sarfChart) {
		this(sarfChart, null);
	}

	/**
	 * @param sarfChart
	 * @param configuration
	 */
	public MainConjugation(SarfChart sarfChart, ChartConfiguration configuration) {
		this.sarfChart = sarfChart;
		this.configuration = configuration == null ? new ChartConfiguration()
				: configuration;

		boolean omitAbbreviatedConjugation = this.sarfChart.getSarfSagheer() == null ? true
				: this.configuration.isOmitAbbreviatedConjugation();
		this.configuration
				.setOmitAbbreviatedConjugation(omitAbbreviatedConjugation);

		boolean omitDetailedConjugation = this.sarfChart.getSarfKabeer() == null ? true
				: this.configuration.isOmitDetailedConjugation();
		this.configuration.setOmitDetailedConjugation(omitDetailedConjugation);

		this.abbreviatedConjugation = new AbbreviatedConjugation(
				this.configuration, this.sarfChart.getSarfSagheer(),
				this.sarfChart.getChartTitle());
		this.detailedConjugation = new DetailedConjugation(
				this.sarfChart.getSarfKabeer());
	}

	public void convert(MainDocumentPart mainDocumentPart)
			throws ApplicationException {
		if (sarfChart == null
				|| (sarfChart.getSarfSagheer() == null && sarfChart
						.getSarfKabeer() == null)) {
			throw new BusinessException(
					"SarfChart is not initailized properly.");
		}

		if (!configuration.isOmitAbbreviatedConjugation()) {
			mainDocumentPart.addObject(abbreviatedConjugation.getChart());
		}
		if (!configuration.isOmitDetailedConjugation()) {
			mainDocumentPart.addObject(detailedConjugation.getChart());
		}
	}

}
