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

import java.net.URI;
import java.net.URISyntaxException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.ExternalRef;
import org.spdx.library.model.ReferenceType;
import org.spdx.library.model.enumerations.ReferenceCategory;
import org.spdx.library.referencetype.ListedReferenceTypes;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;

/**
 * Package external refs
 * @author Gary O'Neall
 *
 */
public class ExternalRefsSheet extends AbstractSheet {
	
	static final Logger logger = LoggerFactory.getLogger(ExternalRefsSheet.class);
	
	static final int PKG_ID_COL = 0;
	static final int REF_CATEGORY_COL = PKG_ID_COL + 1;
	static final int REF_TYPE_COL = REF_CATEGORY_COL + 1;
	static final int REF_LOCATOR_COL = REF_TYPE_COL + 1;
	static final int COMMENT_COL = REF_LOCATOR_COL + 1;

	static final int USER_DEFINED_COLS = COMMENT_COL + 1;
	static final int NUM_COLS = USER_DEFINED_COLS + 1;
	
	static final boolean[] REQUIRED = new boolean[] {true, true, true, true, false, false};
	static final String[] HEADER_TITLES = new String[] {"Package ID", "Category",
		"Type", "Locator", "Comment", "User Defined ..."};
	static final int[] COLUMN_WIDTHS = new int[] {25, 25, 40, 60, 40, 40};
	static final boolean[] LEFT_WRAP = new boolean[] {false, false, true, true, true, true};
	static final boolean[] CENTER_NOWRAP = new boolean[] {true, true, false, false, false, false};

	private static final String NO_REFERENCE_TYPE = "[No Reference Type]";

	/**
	 * @param workbook
	 * @param sheetName
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 */
	public ExternalRefsSheet(Workbook workbook, String sheetName, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, modelStore, documentUri, copyManager);
	}

	/* (non-Javadoc)
	 * @see org.spdx.spdxspreadsheet.AbstractSheet#verify()
	 */
	@Override
	public String verify() {
		try {
			if (sheet == null) {
				return "Worksheet for External Refs does not exist";
			}
			Row firstRow = sheet.getRow(firstRowNum);
			for (int i = 0; i < NUM_COLS- 1; i++) { 	// Don't check the last (user defined) column
				Cell cell = firstRow.getCell(i+firstCellNum);
				if (cell == null || 
						cell.getStringCellValue() == null ||
						!cell.getStringCellValue().equals(HEADER_TITLES[i])) {
					return "Column "+HEADER_TITLES[i]+" missing for External Refs worksheet";
				}
			}
			// validate rows
			boolean done = false;
			int rowNum = getFirstDataRow();
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
			return "Error in verifying External Refs work sheet: "+ex.getMessage();
		}
	}
	
	/**
	 * @param row
	 * @return
	 */
	private String validateRow(Row row) {
		for (int i = 0; i < NUM_COLS; i++) {
			Cell cell = row.getCell(i);
			if (cell == null) {
				if (REQUIRED[i]) {
					return "Required cell "+HEADER_TITLES[i]+" missing for row "+String.valueOf(row.getRowNum());
				}
			}
		}
		return null;
	}

	/**
	 * @param wb
	 * @param externalRefsSheetName
	 */
	public static void create(Workbook wb, String externalRefsSheetName) {
		int sheetNum = wb.getSheetIndex(externalRefsSheetName);
		if (sheetNum >= 0) {
			wb.removeSheetAt(sheetNum);
		}
		Sheet sheet = wb.createSheet(externalRefsSheetName);
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
	 * @param packageId Package ID for the package that contains this external ref
	 * @param externalRef
	 * @throws SpreadsheetException 
	 */
	public void add(String packageId, ExternalRef externalRef) throws SpreadsheetException {
		Row row = addRow();
		if (packageId != null) {
			row.createCell(PKG_ID_COL).setCellValue(packageId);
		}
		try {
			if (externalRef != null) {
				if (externalRef.getReferenceCategory() != null) {
					row.createCell(REF_CATEGORY_COL).setCellValue(externalRef.getReferenceCategory().toString());
				}
				try {
					if (externalRef.getReferenceType() != null) {
						row.createCell(REF_TYPE_COL).setCellValue(refTypeToString(externalRef.getReferenceType()));
					}
				} catch (InvalidSPDXAnalysisException e) {
					throw(new SpreadsheetException("Error getting external reference type: "+e.getMessage()));
				}
				if (externalRef.getReferenceLocator() != null) {
					row.createCell(REF_LOCATOR_COL).setCellValue(externalRef.getReferenceLocator());
				}
				if (externalRef.getComment().isPresent()) {
					row.createCell(COMMENT_COL).setCellValue(externalRef.getComment().get());
				}
			}
		} catch(InvalidSPDXAnalysisException ex) {
			throw new SpreadsheetException("Error getting externalRef from model store",ex);
		}
	}

	/**
	 * Convert a reference type to the type used in 
	 * @param referenceType
	 * @return
	 * @throws InvalidSPDXAnalysisException 
	 */
	protected String refTypeToString(ReferenceType referenceType) {
		String retval;
		if (referenceType == null) {
			return NO_REFERENCE_TYPE;
		}
		String referenceTypeUri = referenceType.getIndividualURI();
		if (ReferenceType.MISSING_REFERENCE_TYPE_URI.equals(referenceTypeUri)) {
			return NO_REFERENCE_TYPE;
		}
		try {
			retval = ListedReferenceTypes.getListedReferenceTypes().getListedReferenceName(new URI(referenceTypeUri));
		} catch (InvalidSPDXAnalysisException e) {
			retval = null;
		} catch (URISyntaxException e) {
			retval = null;
		}
		if (retval == null) {
			retval = referenceTypeUri;
			if (retval.startsWith(documentUri + "#")) {
				retval = retval.substring(documentUri.length()+1);
			}
		}
		return retval;
	}

	/**
	 * Get all external references for a given package ID
	 * @param id
	 * @param container
	 * @return
	 * @throws SpreadsheetException 
	 */
	public List<ExternalRef> getExternalRefsForPkgid(String id) throws SpreadsheetException {
		List<ExternalRef> retval = new ArrayList<>();
		if (id == null || sheet == null) {
			return retval;
		}
		int i = getFirstDataRow();
		Row row = sheet.getRow(i++);
		while(row != null) {
			Cell pkgIdCell = row.getCell(PKG_ID_COL);
			try {
			if (Objects.nonNull(pkgIdCell) && id.equals(pkgIdCell.getStringCellValue())) {
				ExternalRef er = new ExternalRef(modelStore, documentUri, 
						modelStore.getNextId(IdType.Anonymous, documentUri), copyManager, true);
				
				Cell refCategoryCell = row.getCell(REF_CATEGORY_COL);
				if (refCategoryCell != null) {
					try {
						er.setReferenceCategory(ReferenceCategory.valueOf(refCategoryCell.getStringCellValue().trim().replace('-','_')));
					} catch(Exception ex) {
						throw new SpreadsheetException("Invalid reference category: "+refCategoryCell.getStringCellValue());
					}
				}
				
				Cell refTypeCell = row.getCell(REF_TYPE_COL);
				if (refTypeCell != null) {
					String refTypeStr = refTypeCell.getStringCellValue();
					er.setReferenceType(stringToRefType(refTypeStr));
				}
				
				Cell refLocatorCell = row.getCell(REF_LOCATOR_COL);
				if (refLocatorCell != null) {
					er.setReferenceLocator(refLocatorCell.getStringCellValue());
				}
				
				Cell commentCell = row.getCell(COMMENT_COL);
				if (commentCell != null) {
					er.setComment(commentCell.getStringCellValue());
				}
				retval.add(er);
			}
			} catch(InvalidSPDXAnalysisException ex) {
				throw new SpreadsheetException("Error creating ExternalRef",ex);
			}
			row = sheet.getRow(i++);
		}
		return retval;
	}

	/**
	 * Convert a string to a reference type
	 * @param refTypeStr can be a listed reference type name, a URI string, or a local name
	 * @return
	 */
	protected ReferenceType stringToRefType(String refTypeStr) {
		ReferenceType refType = null;
		if (refTypeStr != null) {
			refTypeStr = refTypeStr.trim();
			try {
				refType = ListedReferenceTypes.getListedReferenceTypes().getListedReferenceTypeByName(refTypeStr.trim());
			} catch (InvalidSPDXAnalysisException e) {
				// Ignore - likely due to not being a listed reference type
			}
			if (refType == null) {
				if (!(refTypeStr.contains(":") || refTypeStr.contains("/"))) {
					refTypeStr = documentUri + "#" + refTypeStr;
				}
				try {
					refType = new ReferenceType(refTypeStr);
				} catch (InvalidSPDXAnalysisException e) {
					logger.warn("SPDX Exception creating reference type",e);
				}
			}
		}
		return refType;
	}
}