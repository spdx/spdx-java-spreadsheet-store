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

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxConstants;
import org.spdx.library.model.Annotation;
import org.spdx.library.model.enumerations.AnnotationType;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;

/**
 * Sheet containing all annotations
 * @author Gary O'Neall
 *
 */
public class AnnotationsSheet extends AbstractSheet {

	static final int ID_COL = 0;
	static final int COMMENT_COL = ID_COL + 1;
	static final int DATE_COL = COMMENT_COL + 1;
	static final int ANNOTATOR_COL = DATE_COL + 1;
	static final int TYPE_COL = ANNOTATOR_COL + 1;
	static final int USER_DEFINED_COL = TYPE_COL + 1;
	static final int NUM_COLS = USER_DEFINED_COL;
	
	static final String[] HEADER_TITLES = new String[] {"SPDX Identifier being Annotated",
		"Annotation Comment", "Annotation Date", "Annotator", "Annotation Type",
		"Optional User Defined Columns..."};
	static final int[] COLUMN_WIDTHS = new int[] {25, 70, 25, 60, 20, 50};
	static final boolean[] LEFT_WRAP = new boolean[] {false, true, false, true, false, true};
	static final boolean[] CENTER_NOWRAP = new boolean[] {true, false, true, false, true, false};

	static final boolean[] REQUIRED = new boolean[] {true, true, true, true, true, false};
	
	static final SimpleDateFormat dateFormat = new SimpleDateFormat(SpdxConstants.SPDX_DATE_FORMAT);


	/**
	 * @param workbook Workbook for the sheet
	 * @param annotationsSheetName
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 */
	public AnnotationsSheet(Workbook workbook, String annotationsSheetName, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, annotationsSheetName, modelStore, documentUri, copyManager);
	}

	@Override
	public String verify() {
		try {
			if (sheet == null) {
				return "Worksheet for Annotations does not exist";
			}
			Row firstRow = sheet.getRow(firstRowNum);
			for (int i = 0; i < NUM_COLS; i++) {
				Cell cell = firstRow.getCell(i+firstCellNum);
				if (cell == null || 
						cell.getStringCellValue() == null ||
						!cell.getStringCellValue().equals(HEADER_TITLES[i])) {
					return "Column "+HEADER_TITLES[i]+" missing for Annotation worksheet";
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
			return "Error in verifying Annotations worksheet: "+ex.getMessage();
		}
	}

	private String validateRow(Row row) {
		for (int i = 0; i < NUM_COLS; i++) {
			Cell cell = row.getCell(i);
			if (REQUIRED[i] && cell == null) {
				return "Required cell "+HEADER_TITLES[i]+" missing for row "+String.valueOf(row.getRowNum())+" in annotation sheet";
			} 
			if (i == TYPE_COL && cell.getStringCellValue() != null) {
				try {
					AnnotationType type = AnnotationType.valueOf(cell.getStringCellValue());
					if (type == null) {
						return "Invalid annotation type in row "+String.valueOf(row)+": "+cell.getStringCellValue();
					}
				} catch (Exception ex) {
					return "Invalid annotation type in row "+String.valueOf(row)+": "+cell.getStringCellValue();
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

	/**
	 * @param relationship
	 * @throws SpreadsheetException 
	 */
	public void add(Annotation annotation, String elementId) throws SpreadsheetException {
		Row row = addRow();		
		if (elementId != null) {
			Cell idCell = row.createCell(ID_COL, CellType.STRING);
			idCell.setCellValue(elementId);
		}
		try {
			if (annotation.getComment() != null) {
				row.createCell(COMMENT_COL).setCellValue(annotation.getComment());
			}
			if (annotation.getAnnotationDate() != null) {
				row.createCell(DATE_COL).setCellValue(annotation.getAnnotationDate());
			}
			if (annotation.getAnnotator() != null) {
				row.createCell(ANNOTATOR_COL).setCellValue(annotation.getAnnotator());
			}
			if (annotation.getAnnotationType() != null) {
				row.createCell(TYPE_COL).setCellValue(annotation.getAnnotationType().toString());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting annotation",e);
		}
	}
	
	public String getElmementId(int rowNum) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		return row.getCell(ID_COL).getStringCellValue();
	}
	public Annotation getAnnotation(int rowNum) throws SpreadsheetException {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		String comment = null;
		Cell commentCell = row.getCell(COMMENT_COL);
		if (commentCell != null) {
			comment = commentCell.getStringCellValue();
		} else {
			throw new SpreadsheetException("Missing required annotation comment");
		}
		String date = null;
		Cell dateCell = row.getCell(DATE_COL);
		if (dateCell == null || dateCell.getCellType() == CellType.BLANK) {
		    throw new SpreadsheetException("Missing required annotation date");
		}
		if (dateCell.getCellType() == CellType.STRING) {
			date = dateCell.getStringCellValue();
		} else if (dateCell.getCellType() == CellType.NUMERIC) {
		    date = dateFormat.format(dateCell.getDateCellValue());
		}
		String annotator = null;
		Cell annotatorCell = row.getCell(ANNOTATOR_COL);
		if (annotatorCell != null) {
			annotator = annotatorCell.getStringCellValue();
		} else {
			throw new SpreadsheetException("Missing required annotator");
		}
		AnnotationType type = null;
		Cell typeCell = row.getCell(TYPE_COL);
		if (typeCell != null) {
			try {
				type = AnnotationType.valueOf(typeCell.getStringCellValue().trim());
			} catch(Exception ex) {
				throw new SpreadsheetException("Invalid annotation type");
			}
		} else {
			throw new SpreadsheetException("Missing required spreadsheet type");
		}
		try {
			Annotation retval = new Annotation(modelStore, documentUri, 
					modelStore.getNextId(IdType.Anonymous, documentUri), copyManager, true);
			retval.setAnnotationDate(date);
			retval.setAnnotationType(type);
			retval.setAnnotator(annotator);
			retval.setComment(comment);
			return retval;
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error creating annotation",e);
		}
	}
}