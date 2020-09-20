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

import java.util.Collection;
import java.util.Date;
import java.util.List;
import java.util.Objects;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.ExternalDocumentRef;
import org.spdx.library.model.SpdxDocument;
import org.spdx.storage.IModelStore;

/**
 * Abstract class for sheet containing information about the origins of an SPDX document
 * Specific versions implemented as subclasses
 * @author Gary O'Neall
 *
 */
public abstract class DocumentInfoSheet extends AbstractSheet {
	static final int SPREADSHEET_VERSION_COL = 0;
	static final int DATA_ROW_NUM = 1;
	
	protected String version;

	public DocumentInfoSheet(Workbook workbook, String sheetName, String version, IModelStore modelStore, ModelCopyManager copyManager) throws SpreadsheetException {
		super(workbook, sheetName, modelStore, null, copyManager); // Need to add or create the document from information in the spreadsheet
		this.documentUri = this.getNamespace();
		if (Objects.isNull(this.documentUri)) {
			throw new SpreadsheetException("Missing document URI in the document sheet");
		}
		Objects.requireNonNull(version, "Missing version");
		Objects.requireNonNull(modelStore, "Missing required model store");
		this.version = version;
	}

	/**
	 * @param wb Workbook
	 * @param sheetName Sheet name
	 * @param documentUri Document URI
	 */
	public static void create(Workbook wb, String sheetName, String documentUri) {
		//NOTE: this must be updated to the latest version
		DocumentInfoSheetV2d0.create(wb, sheetName, documentUri);
	}

	/**
	 * Open an existing worksheet
	 * @param workbook
	 * @param originSheetName
	 * @param version Spreadsheet version
	 * @param modelStore model store for the SPDX document
	 * @param copyManager 
	 * @return
	 * @throws SpreadsheetException 
	 */
	public static DocumentInfoSheet openVersion(Workbook workbook,
			String originSheetName, String version, IModelStore modelStore, 
			ModelCopyManager copyManager) throws SpreadsheetException {
		return new DocumentInfoSheetV2d0(workbook, originSheetName, version, modelStore, copyManager);
	}
	
	protected Row getDataRow() {
		return getDataRow(0);
	}
	
	protected Row getDataRow(int rowIndex) {
		while (firstRowNum + DATA_ROW_NUM + rowIndex > lastRowNum) {
			addRow();
		}
		Row dataRow = sheet.getRow(firstRowNum + DATA_ROW_NUM + rowIndex);
		if (dataRow == null) {
			dataRow = sheet.createRow(firstRowNum + DATA_ROW_NUM + rowIndex);
		}
		return dataRow;
	}
	
	protected Cell getOrCreateDataCell(int colNum) {
		Cell cell = getDataRow().getCell(colNum);
		if (cell == null) {
			cell = getDataRow().createCell(colNum);
			//cell.setCellType(CellType.NUMERIC);
		}
		return cell;
	}
	
	protected void setDataCellStringValue(int colNum, String value) {
		getOrCreateDataCell(colNum).setCellValue(value);
	}
	
	protected void setDataCellDateValue(int colNum, Date value) {
		Cell cell = getOrCreateDataCell(colNum);
		cell.setCellValue(value);
		cell.setCellStyle(dateStyle);
		
	}
	
	protected Date getDataCellDateValue(int colNum) {
		Cell cell = getDataRow().getCell(colNum);
		if (cell == null) {
			return null;
		} else {
			return cell.getDateCellValue();
		}
	}

	protected String getDataCellStringValue(int colNum) {
		Cell cell = getDataRow().getCell(colNum);
		if (cell == null) {
			return null;
		} else {
			if (cell.getCellType() == CellType.NUMERIC) {				
				return Double.toString(cell.getNumericCellValue());
			} else {
				return cell.getStringCellValue();
			}
		}
	}

	/**
	 * @param spdxVersion
	 */
	public abstract void setSPDXVersion(String spdxVersion);

	/**
	 * @param createdBys
	 */
	public abstract void setCreatedBy(Collection<String> createdBys);

	/**
	 * @param id
	 */
	public abstract void setDataLicense(String licenseId);

	/**
	 * @param comments
	 */
	public abstract void setAuthorComments(String comments);

	/**
	 * @param parse
	 */
	public abstract void setCreated(Date createdDate);

	/**
	 * @return
	 */
	public abstract Date getCreated();

	/**
	 * @return
	 */
	public abstract List<String> getCreatedBy();

	/**
	 * @return
	 */
	public abstract String getAuthorComments();

	/**
	 * @return
	 */
	public abstract String getSPDXVersion();

	/**
	 * @return
	 */
	public abstract String getDataLicense();

	/**
	 * @return
	 */
	public abstract String getDocumentComment();

	/**
	 * @param docComment
	 */
	public abstract void setDocumentComment(String docComment);
	
	/**
	 * @return
	 */
	public abstract String getLicenseListVersion();
	
	/**
	 * @param licenseVersion
	 */
	public abstract void setLicenseListVersion(String licenseVersion);

	/**
	 * @return
	 */
	public abstract String getNamespace();

	/**
	 * Add all origin information from the document
	 * @param doc
	 * @throws SpreadsheetException 
	 */
	public abstract void addDocument(SpdxDocument doc) throws SpreadsheetException;
	
	/**
	 * @return SPDX Identifier for the document
	 */
	public abstract String getSpdxId();
	/**
	 * Set the SPDX identified for the document
	 * @param id
	 */
	public abstract void setSpdxId(String id);
	/**
	 * @return Document name
	 */
	public abstract String getDocumentName();
	/**
	 * Set the document name
	 * @param documentName
	 */
	public abstract void setDocumentName(String documentName);
	/**
	 * @return SPDX ID's for content described by this SPDX document
	 */
	public abstract Collection<String> getDocumentContents();
	/**
	 * Set the SPDX ID's for content described by this SPDX document
	 * @param contents
	 */
	public abstract void setDocumentDescribes(Collection<String> contents);
	/**
	 * @return External document refs
	 * @throws SpreadsheetException 
	 */
	public abstract Collection<ExternalDocumentRef> getExternalDocumentRefs() throws SpreadsheetException;
	/**
	 * Set the external document refs
	 * @param externalDocumentRefs
	 * @throws SpreadsheetException 
	 */
	public abstract void setExternalDocumentRefs(Collection<ExternalDocumentRef> externalDocumentRefs) throws SpreadsheetException;
}
