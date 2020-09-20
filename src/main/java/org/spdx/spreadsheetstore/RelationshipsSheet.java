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

import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.ModelObject;
import org.spdx.library.model.Relationship;
import org.spdx.library.model.SpdxElement;
import org.spdx.library.model.SpdxModelFactory;
import org.spdx.library.model.enumerations.RelationshipType;
import org.spdx.storage.IModelStore;

/**
 * Sheet containing relationship data
 * @author Gary O'Neall
 *
 */
public class RelationshipsSheet extends AbstractSheet {

	static final int ID_COL = 0;
	static final int RELATIONSHIP_COL = ID_COL + 1;
	static final int RELATED_ID_COL = RELATIONSHIP_COL + 1;
	static final int COMMENT_COL = RELATED_ID_COL + 1;
	static final int USER_DEFINED_COL = COMMENT_COL + 1;
	static final int NUM_COLS = USER_DEFINED_COL;
	
	static final String[] HEADER_TITLES = new String[] {"SPDX Identifier A",
		"Relationship", "SPDX Identifier B", "Relationship Comment", 
		"Optional User Defined Columns..."};
	static final int[] COLUMN_WIDTHS = new int[] {20, 25, 20, 70, 50};
	static final boolean[] LEFT_WRAP = new boolean[] {false, false, false, true, true};
	static final boolean[] CENTER_NOWRAP = new boolean[] {true, true, true, false, false};

	static final boolean[] REQUIRED = new boolean[] {true, true, true, false, false};


	/**
	 * @param workbook
	 * @param sheetName
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 */
	public RelationshipsSheet(Workbook workbook, String sheetName, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, modelStore, documentUri, copyManager);
	}

	@Override
	public String verify() {
		try {
			if (sheet == null) {
				return "Worksheet for Relationships does not exist";
			}
			Row firstRow = sheet.getRow(firstRowNum);
			for (int i = 0; i < NUM_COLS; i++) {
				Cell cell = firstRow.getCell(i+firstCellNum);
				if (cell == null || 
						cell.getStringCellValue() == null ||
						!cell.getStringCellValue().equals(HEADER_TITLES[i])) {
					return "Column "+HEADER_TITLES[i]+" missing for Relationship worksheet";
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
			return "Error in verifying Relationship worksheet: "+ex.getMessage();
		}
	}

	private String validateRow(Row row) {
		for (int i = 0; i < NUM_COLS; i++) {
			Cell cell = row.getCell(i);
			if (REQUIRED[i] && cell == null) {
				return "Required cell "+HEADER_TITLES[i]+" missing for row "+String.valueOf(row.getRowNum())+" in relationships sheet";
			} 
			if (i == RELATIONSHIP_COL && cell.getStringCellValue() != null) {
				try {
					RelationshipType.valueOf(cell.getStringCellValue().trim());
				} catch (IllegalArgumentException ex) {
					return "Invalid relationship type in row "+String.valueOf(row) + ": " + cell.getStringCellValue();
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
	public void add(Relationship relationship, String elementId) throws SpreadsheetException {
		Row row = addRow();		
		if (elementId != null) {
			Cell idCell = row.createCell(ID_COL, CellType.STRING);
			idCell.setCellValue(elementId);
		}
		try {
			if (relationship.getRelationshipType() != null) {
				Cell relationshipCell = row.createCell(RELATIONSHIP_COL, CellType.STRING);
				relationshipCell.setCellValue(relationship.getRelationshipType().toString());
			}
			if (relationship.getRelatedSpdxElement().isPresent()) {
				Cell relatedIdCell = row.createCell(RELATED_ID_COL, CellType.STRING);
				relatedIdCell.setCellValue(relationship.getRelatedSpdxElement().get().getId());
			}		
			if (relationship.getComment() .isPresent()) {
				Cell commentCell = row.createCell(COMMENT_COL, CellType.STRING);
				commentCell.setCellValue(relationship.getComment().get());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting relationship fields",e);
		}
	}
	
	public String getElmementId(int rowNum) {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		return row.getCell(ID_COL).getStringCellValue();
	}
	
	public Relationship getRelationship(int rowNum) throws SpreadsheetException {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		Cell relatedIdCell = row.getCell(RELATED_ID_COL);
		String relatedId = null;
		if (relatedIdCell != null && relatedIdCell.getStringCellValue() != null) {
			relatedId = relatedIdCell.getStringCellValue();
		}
		RelationshipType type = null;
		Cell relationshipCell = row.getCell(RELATIONSHIP_COL);
		if (relationshipCell != null && relationshipCell.getStringCellValue() != null) {
			try {
				type = RelationshipType.valueOf(relationshipCell.getStringCellValue().trim());
			} catch(Exception ex) {
				throw new SpreadsheetException("Invalid relationship type: "+relationshipCell.getStringCellValue().trim());
			}
		}
		Cell commentCell = row.getCell(COMMENT_COL);
		String comment = null;
		if (commentCell != null && commentCell.getStringCellValue() != null) {
			comment = commentCell.getStringCellValue();
		}
		if (relatedId == null) {
			throw new SpreadsheetException("No related element ID for relationship");
		}
		SpdxElement element;
		try {
			Optional<ModelObject> mo = SpdxModelFactory.getModelObject(modelStore, documentUri, relatedId, copyManager);
			if (!mo.isPresent()) {
				throw new SpreadsheetException("No element found for relationship with related ID "+relatedId);
			}
			if (!(mo.get() instanceof SpdxElement)) {
				throw new SpreadsheetException("Related ID "+relatedId+" is not of type SpdxElement");
			}
			element = (SpdxElement)mo.get();
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting SpdxElement for relationship with related ID "+relatedId,e);
		}	
		try {
			return element.createRelationship(element, type, comment);
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error creating relationship for related ID "+relatedId,e);
		}
	}
}