/**
 * Copyright (c) 2020 Source Auditor Inc.
 *
 * SPDX-License-Identifier: Apache-2.0
 * 
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */
package org.spdx.spreadsheetstore;

import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxConstants;
import org.spdx.library.model.SpdxDocument;
import org.spdx.library.model.enumerations.AnnotationType;
import org.spdx.storage.IModelStore;

/**
 * Sheet describing any reviewers for an SPDX Document
 * @author Gary O'Neall
 *
 */
public class ReviewersSheet extends AbstractSheet {

	static final int NUM_COLS = 3;
	static final int REVIEWER_COL = 0;
	static final int TIMESTAMP_COL = REVIEWER_COL + 1;
	static final int COMMENT_COL = TIMESTAMP_COL + 1;
	static final String[] HEADER_TITLES = new String[] {"Reviewer", "Review Date", "Reviewer Comment"};
	static final int[] COLUMN_WIDTHS = new int[] {60, 20, 120};
	static final boolean[] LEFT_WRAP = new boolean[] {true, false, true};
	static final boolean[] CENTER_NOWRAP = new boolean[] {false, true, false};

	static final boolean[] REQUIRED = new boolean[] {true, true, false};
	
	SimpleDateFormat spdxFormat = new SimpleDateFormat(SpdxConstants.SPDX_DATE_FORMAT);
	
	public ReviewersSheet(Workbook workbook, String sheetName, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, modelStore, documentUri, copyManager);
	}

	@Override
	public String verify() {
		try {
			if (sheet == null) {
				return "Worksheet for SPDX Reviewers does not exist";
			}
			Row firstRow = sheet.getRow(firstRowNum);
			for (int i = 0; i < NUM_COLS; i++) {
				Cell cell = firstRow.getCell(i+firstCellNum);
				if (cell == null || 
						cell.getStringCellValue() == null ||
						!cell.getStringCellValue().equals(HEADER_TITLES[i])) {
					return "Column "+HEADER_TITLES[i]+" missing for SPDX Reviewers worksheet";
				}
			}
			// validate rows
			boolean done = false;
			int rowNum = firstRowNum + 1;
			while (!done) {
				Row row = sheet.getRow(rowNum);
				if (row == null || row.getCell(firstCellNum) == null) {
					done = true;
				} else {
					String error = validateRow(row);
					if (error != null) {
						return error;
					}
					rowNum++;
				}
			}
			return null;
		} catch (Exception ex) {
			return "Error in verifying SPDX Reviewers worksheet: "+ex.getMessage();
		}
	}

	private String validateRow(Row row) {
		for (int i = 0; i < NUM_COLS; i++) {
			Cell cell = row.getCell(i);
			if (REQUIRED[i] && cell == null) {
				return "Required cell "+HEADER_TITLES[i]+" missing for row "+String.valueOf(row.getRowNum())+" in reviewer sheet";
			} else {
				if (i == TIMESTAMP_COL) {
					if (!(cell.getCellType() == CellType.NUMERIC)) {
						return "Timestamp cell is not a numeric type for row "+String.valueOf(row.getRowNum())+" in Reviewer sheet";
					}
				}
			}
		}
		return null;
	}

	public static void create(Workbook wb, String sheetName) {
		int sheetNum = wb.getSheetIndex(sheetName);
		if (sheetNum >= 0) {
			wb.removeSheetAt(sheetNum);
		}
		Sheet sheet = wb.createSheet(sheetName);
		CellStyle headerStyle = AbstractSheet.createHeaderStyle(wb);
		CellStyle centerStyle = AbstractSheet.createCenterStyle(wb);
		CellStyle wrapStyle = AbstractSheet.createLeftWrapStyle(wb);
		Row row = sheet.createRow(0);
		for (int i = 0; i < HEADER_TITLES.length; i++) {
			sheet.setColumnWidth(i, COLUMN_WIDTHS[i]*256);
			if (LEFT_WRAP[i]) {
				sheet.setDefaultColumnStyle(i, wrapStyle);
			} else if (CENTER_NOWRAP[i]) {
				sheet.setDefaultColumnStyle(i, centerStyle);
			}
			Cell cell = row.createCell(i);
			cell.setCellStyle(headerStyle);
			cell.setCellValue(HEADER_TITLES[i]);
		}
	}

	public String getReviewer(int rowNum) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		Cell reviewer = row.getCell(REVIEWER_COL);
		if (reviewer == null) {
			return null;
		}
		return reviewer.getStringCellValue();
	}
	
	public Date getReviewerTimestamp(int rowNum) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		Cell tsCell = row.getCell(TIMESTAMP_COL);
		if (tsCell == null) {
			return null;
		}
		return tsCell.getDateCellValue();
	}
	
	public String getReviewerComment(int rowNum) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		Cell commentCell = row.getCell(COMMENT_COL);
		if (commentCell == null) {
			return null;
		}
		return commentCell.getStringCellValue();
	}
	
	/**
	 * Add all of the reviews as REVIEW annotations to the SPDX document
	 * @throws SpreadsheetException
	 */
	public void addReviewsToDocAnnotations() throws SpreadsheetException {
		SpdxDocument doc;
		try {
			doc = new SpdxDocument(modelStore, documentUri, copyManager, false);
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("SPDX document not found.  Create the SPDX document before adding reviewers.",e);
		}
		int rowNum = getFirstDataRow();
		Row row = sheet.getRow(rowNum++);
		while (Objects.nonNull(row) && Objects.nonNull(row.getCell(REVIEWER_COL)) && 
				!row.getCell(REVIEWER_COL).getStringCellValue().isEmpty()) {
			String annotator = row.getCell(REVIEWER_COL).getStringCellValue();
			String date;
			Cell dateCell = row.getCell(TIMESTAMP_COL);
			if (Objects.nonNull(dateCell)) {
				date = spdxFormat.format(dateCell.getDateCellValue());
			} else {
				date = spdxFormat.format(new Date());
			}
			Cell commentCell = row.getCell(COMMENT_COL);
			String comment;
			if (Objects.nonNull(commentCell)) {
				comment = commentCell.getStringCellValue();
			} else {
				comment = "";
			}
			try {
				doc.addAnnotation(doc.createAnnotation(annotator, AnnotationType.REVIEW, date, comment));
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error adding annotation",e);
			}
			row = sheet.getRow(rowNum++);
		}
	}
}
