/*
 * SPDX-FileContributor: Gary O'Neall
 * SPDX-FileCopyrightText: Copyright (c) 2020 Source Auditor Inc.
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 * <p>
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 * <p>
 *       https://www.apache.org/licenses/LICENSE-2.0
 * <p>
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
import org.spdx.library.model.v2.ExternalDocumentRef;
import org.spdx.library.model.v2.SpdxDocument;
import org.spdx.storage.IModelStore;

/**
 * Abstract class for sheet containing information about the origins of an SPDX document
 *
 * Specific versions implemented as subclasses
 * @author Gary O'Neall
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
	 * Construct a DocumentInfoSheet in the given workbook
	 *
	 * @param wb The Workbook object where the sheet will be created
	 * @param sheetName Sheet name
	 * @param documentUri Document URI
	 */
	public static void create(Workbook wb, String sheetName, String documentUri) {
		//NOTE: this must be updated to the latest version
		DocumentInfoSheetV2d0.create(wb, sheetName, documentUri);
	}

	/**
	 * Open an existing worksheet
	 *
	 * @param workbook Workbook object containing the sheet.
	 * @param originSheetName Name of the sheet to be opened.
	 * @param version Spreadsheet version
	 * @param modelStore Model store for the SPDX document
	 * @param copyManager The copy manager used for handling SPDX object copies.
	 * @return An instance of {@link DocumentInfoSheet} corresponding to the specified version.
	 * @throws SpreadsheetException If the sheet cannot be opened or the version is unsupported.
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
	 * Se the SPDX specification version for the SPDX document
	 *
	 * @param spdxVersion SPDX specification version
	 */
	public abstract void setSPDXVersion(String spdxVersion);

	/**
	 * Set the list of creators of the SPDX document
	 *
	 * @param createdBys List of creators
	 */
	public abstract void setCreatedBy(Collection<String> createdBys);

	/**
	 * Set the license for the SPDX document
	 *
	 * @param licenseId License ID
	 */
	public abstract void setDataLicense(String licenseId);

	/**
	 * Set the author comments for the SPDX document
	 *
	 * @param comments Author comments
	 */
	public abstract void setAuthorComments(String comments);

	/**
	 * Set the created date for the SPDX document
	 *
	 * @param createdDate Created date
	 */
	public abstract void setCreated(Date createdDate);

	/**
	 * Retrieve the creation date of the SPDX document
	 *
	 * @return Date the SPDX document was created
	 */
	public abstract Date getCreated();

	/**
	 * Retrieve the list of creators of the SPDX document
	 *
	 * @return List of creators
	 */
	public abstract List<String> getCreatedBy();

	/**
	 * Retrieve the author comments for the SPDX document
	 *
	 * @return Author comments
	 */
	public abstract String getAuthorComments();

	/**
	 * Retrieve the SPDX specification version of the SPDX document
	 *
	 * @return SPDX specification version
	 */
	public abstract String getSPDXVersion();

	/**
	 * Retrieve the license for the SPDX document
	 *
	 * @return License ID
	 */
	public abstract String getDataLicense();

	/**
	 * Retrieve the document comment for the SPDX document
	 *
	 * @return Document comment
	 */
	public abstract String getDocumentComment();

	/**
	 * Set the document comment for the SPDX document
	 *
	 * @param docComment Document comment
	 */
	public abstract void setDocumentComment(String docComment);
	
	/**
	 * Retrieve the license list version for the SPDX document
	 *
	 * @return License list version
	 */
	public abstract String getLicenseListVersion();
	
	/**
	 * Set the license list version for the SPDX document
	 *
	 * @param licenseVersion License list version
	 */
	public abstract void setLicenseListVersion(String licenseVersion);

	/**
	 * Retrieve the namespace for the SPDX document
	 *
	 * @return Namespace
	 */
	public abstract String getNamespace();

	/**
	 * Add all origin information from the document
	 *
	 * @param doc SPDX document to add
	 * @throws SpreadsheetException 
	 */
	public abstract void addDocument(SpdxDocument doc) throws SpreadsheetException;
	
	/**
	 * Retrieve the SPDX identifier for the SPDX document
	 *
	 * @return SPDX identifier for the document
	 */
	public abstract String getSpdxId();

	/**
	 * Set the SPDX identifier for the document
	 *
	 * @param id SPDX identifier
	 */
	public abstract void setSpdxId(String id);

	/**
	 * Retrieve the document name
	 *
	 * @return Document name
	 */
	public abstract String getDocumentName();

	/**
	 * Set the document name
	 *
	 * @param documentName Document name
	 */
	public abstract void setDocumentName(String documentName);

	/**
	 * @return SPDX ID's for content described by this SPDX document
	 */
	public abstract Collection<String> getDocumentContents();

	/**
	 * Set the SPDX ID's for content described by this SPDX document
	 *
	 * @param contents SPDX ID's for content described by this SPDX document
	 */
	public abstract void setDocumentDescribes(Collection<String> contents);

	/**
	 * Retrieve the external document refs
	 *
	 * @return External document refs
	 * @throws SpreadsheetException 
	 */
	public abstract Collection<ExternalDocumentRef> getExternalDocumentRefs() throws SpreadsheetException;

	/**
	 * Set the external document refs
	 *
	 * @param externalDocumentRefs
	 * @throws SpreadsheetException 
	 */
	public abstract void setExternalDocumentRefs(Collection<ExternalDocumentRef> externalDocumentRefs) throws SpreadsheetException;
}
