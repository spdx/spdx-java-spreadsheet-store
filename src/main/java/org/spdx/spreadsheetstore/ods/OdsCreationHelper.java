/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.ExtendedColor;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

/**
 * Adapter for Apache POI {@link CreationHelper} for ODS documents.
 * Provides factories for formula evaluators, formats, and hyperlinks.
 */
public class OdsCreationHelper implements CreationHelper {

	private final OdsWorkbook workbook;
	private final OdsDataFormat dataFormat = new OdsDataFormat();

	/**
	 * Creates an ODS creation helper adapter.
	 *
	 * @param workbook Parent ODS workbook.
	 */
	public OdsCreationHelper(OdsWorkbook workbook) {
		this.workbook = workbook;
	}

	@Override
	public FormulaEvaluator createFormulaEvaluator() {
		return null;
	}

	@Override
	public OdsDataFormat createDataFormat() {
		return dataFormat;
	}

	@Override
	public Hyperlink createHyperlink(HyperlinkType type) {
		return null;
	}

	@Override
	public ExtendedColor createExtendedColor() {
		return null;
	}

	@Override
	public ClientAnchor createClientAnchor() {
		return null;
	}

	@Override
	public AreaReference createAreaReference(String reference) {
		return new AreaReference(reference, workbook.getSpreadsheetVersion());
	}

	@Override
	public AreaReference createAreaReference(CellReference topLeft, CellReference bottomRight) {
		return new AreaReference(topLeft, bottomRight, workbook.getSpreadsheetVersion());
	}

	@Override
	public RichTextString createRichTextString(String text) {
		return new OdsRichTextString(text);
	}
}
