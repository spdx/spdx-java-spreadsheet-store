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

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Objects;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.spdx.library.ModelCopyManager;
import org.spdx.storage.IModelStore;

/**
 * Spreadsheet workbook for an SPDX Document
 * 
 * @author Gary O'Neall
 *
 */
public class SpdxSpreadsheet {
	
	static final Logger logger = LoggerFactory.getLogger(SpdxSpreadsheet.class);
	
	/*
	 * The following information relates to the version management for the SPDXSpreadsheet.
	 * Each sheet in the workbook implements a Factory method to instantiate the correct
	 * version using the static method <code>openVersion(Workbook wb, String sheetName, String versionNumber)</code>
	 * Each sheet also implements a method to create the latest version <code>create(Workbook wb, String sheetName)</code>
	 */
	public static final String CURRENT_VERSION = "2.2.0";
	public static final String VERSION_2_1_0 = "2.1.0";
	public static final String VERSION_2_0_0 = "2.0.0";

	public static final String UNKNOWN_VERSION = "UNKNOWN";
	public static final String UNKNOWN_NAMESPACE = "http://spdx.unknown.namespace";
	public static final List<String> SUPPORTED_VERSIONS = Collections.unmodifiableList(Arrays.asList(
			new String[]{CURRENT_VERSION, VERSION_2_1_0, VERSION_2_0_0}));
	
	Workbook workbook;
	
	private DocumentInfoSheet documentInfoSheet;
	static final String DOCUMENT_INFO_NAME = "Document Info";
	private PackageInfoSheet packageInfoSheet;
	static final String PACKAGE_INFO_SHEET_NAME = "Package Info";
	private ExtractedLicenseInfoSheet extractedLicenseInfoSheet;
	static final String NON_STANDARD_LICENSE_SHEET_NAME = "Extracted License Info";
	private PerFileSheet perFileSheet;
	static final String PER_FILE_SHEET_NAME = "Per File Info";
	private RelationshipsSheet relationshipsSheet;
	static final String RELATIONSHIPS_SHEET_NAME = "Relationships";
	private AnnotationsSheet annotationsSheet;
	static final String ANNOTATIONS_SHEET_NAME = "Annotations";
	private ReviewersSheet reviewersSheet;
	static final String REVIEWERS_SHEET_NAME = "Reviewers";
	private SnippetSheet snippetSheet;
	static final String SNIPPET_SHEET_NAME = "Snippets";
	private ExternalRefsSheet externalRefsSheet;
	static final String EXTERNAL_REFS_SHEET_NAME = "External Refs";

	private IModelStore modelStore;
	private String documentUri;
	private String version;

	private ModelCopyManager copyManager;

	/**
	 * Open an existing SPDX spreadsheet from an input stream
	 * @param stream
	 * @param modelStore
	 * @param copyManager
	 * @throws SpreadsheetException
	 */
	public SpdxSpreadsheet(InputStream stream, IModelStore modelStore, ModelCopyManager copyManager) throws SpreadsheetException {
		Objects.requireNonNull(modelStore, "Missing required model store");
		Objects.requireNonNull(copyManager, "Missing required model copy manager");
		this.modelStore = modelStore;
		this.copyManager = copyManager;
		try {
			workbook = WorkbookFactory.create(stream);
		} catch (EncryptedDocumentException e) {
			logger.error("Unable to read encrypted SPDX Spreadsheet", e);
			throw new SpreadsheetException("Unable to read encrypted SPDX Spreadsheet", e);
		} catch (IOException e) {
			logger.error("I/O error reading SPDX Spreadsheet", e);
			throw new SpreadsheetException("I/O error reading SPDX Spreadsheet", e);
		}
		this.version = readVersion(this.workbook, DOCUMENT_INFO_NAME);
		if (this.version.equals(UNKNOWN_VERSION)) {
			throw(new SpreadsheetException("The version for the SPDX spreadsheet could not be read."));
		}
		this.documentInfoSheet = DocumentInfoSheet.openVersion(this.workbook, DOCUMENT_INFO_NAME, this.version, modelStore, copyManager);
		String verifyMsg = documentInfoSheet.verify();
		if (verifyMsg != null) {
			logger.error(verifyMsg);
			throw(new SpreadsheetException(verifyMsg));
		}
		this.documentUri = this.documentInfoSheet.getNamespace();
		this.packageInfoSheet = PackageInfoSheet.openVersion(this.workbook, PACKAGE_INFO_SHEET_NAME, this.version, modelStore, this.documentUri, copyManager);
		this.extractedLicenseInfoSheet = ExtractedLicenseInfoSheet.openVersion(this.workbook, NON_STANDARD_LICENSE_SHEET_NAME, version, modelStore, this.documentUri, copyManager);
		this.perFileSheet = PerFileSheet.openVersion(this.workbook, PER_FILE_SHEET_NAME, version, modelStore, this.documentUri, copyManager);
		this.relationshipsSheet = new RelationshipsSheet(this.workbook, RELATIONSHIPS_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.annotationsSheet = new AnnotationsSheet(this.workbook, ANNOTATIONS_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.reviewersSheet = new ReviewersSheet(this.workbook, REVIEWERS_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.snippetSheet = new SnippetSheet(this.workbook, SNIPPET_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.externalRefsSheet = new ExternalRefsSheet(this.workbook, EXTERNAL_REFS_SHEET_NAME, modelStore, this.documentUri, copyManager);

		verifyMsg = verifyWorkbook();
		if (verifyMsg != null) {
			logger.error(verifyMsg);
			throw(new SpreadsheetException(verifyMsg));
		}
	}
	
	/**
	 * Create a blank SPDX spreadsheet
	 * @param modelStore
	 * @param copyManager
	 * @throws SpreadsheetException 
	 */
	public SpdxSpreadsheet(IModelStore modelStore, ModelCopyManager copyManager, String documentUri) throws SpreadsheetException {
		Objects.requireNonNull(modelStore, "Missing required model store");
		Objects.requireNonNull(copyManager, "Missing required model copy manager");
		this.modelStore = modelStore;
		this.copyManager = copyManager;
		this.version = CURRENT_VERSION;
		workbook = new XSSFWorkbook();
		this.documentUri = documentUri;
		create();
		this.documentInfoSheet = DocumentInfoSheet.openVersion(this.workbook, DOCUMENT_INFO_NAME, this.version, modelStore, copyManager);
		this.packageInfoSheet = PackageInfoSheet.openVersion(this.workbook, PACKAGE_INFO_SHEET_NAME, this.version, modelStore, this.documentUri, copyManager);
		this.extractedLicenseInfoSheet = ExtractedLicenseInfoSheet.openVersion(this.workbook, NON_STANDARD_LICENSE_SHEET_NAME, version, modelStore, this.documentUri, copyManager);
		this.perFileSheet = PerFileSheet.openVersion(this.workbook, PER_FILE_SHEET_NAME, version, modelStore, this.documentUri, copyManager);
		this.relationshipsSheet = new RelationshipsSheet(this.workbook, RELATIONSHIPS_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.annotationsSheet = new AnnotationsSheet(this.workbook, ANNOTATIONS_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.reviewersSheet = new ReviewersSheet(this.workbook, REVIEWERS_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.snippetSheet = new SnippetSheet(this.workbook, SNIPPET_SHEET_NAME, modelStore, this.documentUri, copyManager);
		this.externalRefsSheet = new ExternalRefsSheet(this.workbook, EXTERNAL_REFS_SHEET_NAME, modelStore, this.documentUri, copyManager);
	}
	
	private void create() throws SpreadsheetException {
		DocumentInfoSheet.create(workbook, DOCUMENT_INFO_NAME, documentUri);
		PackageInfoSheet.create(workbook, PACKAGE_INFO_SHEET_NAME);
		ExternalRefsSheet.create(workbook, EXTERNAL_REFS_SHEET_NAME);
		ExtractedLicenseInfoSheet.create(workbook, NON_STANDARD_LICENSE_SHEET_NAME);
		PerFileSheet.create(workbook, PER_FILE_SHEET_NAME);
		RelationshipsSheet.create(workbook, RELATIONSHIPS_SHEET_NAME);
		AnnotationsSheet.create(workbook, ANNOTATIONS_SHEET_NAME);
		SnippetSheet.create(workbook, SNIPPET_SHEET_NAME);
		ReviewersSheet.create(workbook, REVIEWERS_SHEET_NAME);
	}
	
	public void clear() {
		this.documentInfoSheet.clear();
		this.packageInfoSheet.clear();
		this.extractedLicenseInfoSheet.clear();
		this.perFileSheet.clear();
		this.relationshipsSheet.clear();
		this.annotationsSheet.clear();
		this.reviewersSheet.clear();
		this.snippetSheet.clear();
		this.externalRefsSheet.clear();
	}
	
	/**
	 * Determine the version of an existing workbook
	 * @param workbook
	 * @param originSheetName
	 * @return
	 * @throws SpreadsheetException 
	 */
	private String readVersion(Workbook workbook, String originSheetName) throws SpreadsheetException {
		Sheet sheet = workbook.getSheet(originSheetName);
		if (sheet == null) {
			throw new SpreadsheetException("Invalid SPDX spreadsheet.  Sheet "+originSheetName+" does not exist.");
		}
		int firstRowNum = sheet.getFirstRowNum();
		Row dataRow = sheet.getRow(firstRowNum + DocumentInfoSheet.DATA_ROW_NUM);
		if (dataRow == null) {
			return UNKNOWN_VERSION;
		}
		Cell versionCell = dataRow.getCell(DocumentInfoSheet.SPREADSHEET_VERSION_COL);
		if (versionCell == null) {
			return UNKNOWN_VERSION;
		}
		return versionCell.getStringCellValue();
	}
	
	public static boolean verifyVersion(String ver) {
		return SUPPORTED_VERSIONS.contains(ver);
	}
	
	public String verifyWorkbook() {
		String retval = this.documentInfoSheet.verify();
		if (retval == null || retval.isEmpty()) {
			retval = this.packageInfoSheet.verify();
		}
		if (retval == null || retval.isEmpty()) {
			retval = this.extractedLicenseInfoSheet.verify();
		}
		if (retval == null || retval.isEmpty()) {
			retval = this.perFileSheet.verify();
		}
		if (retval == null || retval.isEmpty()) {
			retval = this.reviewersSheet.verify();
		}
		if (retval == null || retval.isEmpty()) {
			retval = this.relationshipsSheet.verify();
		}
		if (retval == null || retval.isEmpty()) {
			retval = this.annotationsSheet.verify();
		}
		if ((retval == null || retval.isEmpty()) && VERSION_2_0_0.compareTo(this.version) < 0) {
			retval = this.snippetSheet.verify();
		}
		if ((retval == null || retval.isEmpty()) && VERSION_2_0_0.compareTo(this.version) < 0) {
			retval = this.externalRefsSheet.verify();
		}
		
		return retval;
	}

	/**
	 * @return the documentUri
	 */
	public String getDocumentUri() {
		return this.documentUri;
	}
	
	/**
	 * @return the originsSheet
	 */
	public DocumentInfoSheet getOriginsSheet() {
		return documentInfoSheet;
	}

	/**
	 * @param originsSheet the originsSheet to set
	 */
	public void setOriginsSheet(DocumentInfoSheet originsSheet) {
		this.documentInfoSheet = originsSheet;
	}

	/**
	 * @return the packageInfoSheet
	 */
	public PackageInfoSheet getPackageInfoSheet() {
		return packageInfoSheet;
	}

	/**
	 * @return the perFileSheet
	 */
	public PerFileSheet getPerFileSheet() {
		return perFileSheet;
	}

	/**
	 * @return the reviewersSheet
	 */
	public ReviewersSheet getReviewersSheet() {
		return reviewersSheet;
	}

	/**
	 * @param reviewersSheet the reviewersSheet to set
	 */
	public void setReviewersSheet(ReviewersSheet reviewersSheet) {
		this.reviewersSheet = reviewersSheet;
	}
	
	public RelationshipsSheet getRelationshipsSheet() {
		return relationshipsSheet;
	}

	public void setRelationshipsSheet(RelationshipsSheet relationshipsSheet) {
		this.relationshipsSheet = relationshipsSheet;
	}

	public AnnotationsSheet getAnnotationsSheet() {
		return annotationsSheet;
	}

	public void setAnnotationsSheet(AnnotationsSheet annotationsSheet) {
		this.annotationsSheet = annotationsSheet;
	}

	public void setPackageInfoSheet(PackageInfoSheet packageInfoSheet) {
		this.packageInfoSheet = packageInfoSheet;
	}

	public void setPerFileSheet(PerFileSheet perFileSheet) {
		this.perFileSheet = perFileSheet;
	}
	
	/**
	 * @return the snippetSheet
	 */
	public SnippetSheet getSnippetSheet() {
		return snippetSheet;
	}

	/**
	 * @param snippetSheet the snippetSheet to set
	 */
	public void setSnippetSheet(SnippetSheet snippetSheet) {
		this.snippetSheet = snippetSheet;
	}
	
	/**
	 * @return the externalRefsSheet
	 */
	public ExternalRefsSheet getExternalRefsSheet() {
		return externalRefsSheet;
	}

	/**
	 * @param snippetSheet the snippetSheet to set
	 */
	public void setExternaRefsSheet(ExternalRefsSheet externalRefsSheet) {
		this.externalRefsSheet = externalRefsSheet;
	}

	
	/**
	 * @return the documentInfoSheet
	 */
	public DocumentInfoSheet getDocumentInfoSheet() {
		return documentInfoSheet;
	}

	/**
	 * @param documentInfoSheet the documentInfoSheet to set
	 */
	public void setDocumentInfoSheet(DocumentInfoSheet documentInfoSheet) {
		this.documentInfoSheet = documentInfoSheet;
	}

	/**
	 * @return the extractedLicenseInfoSheet
	 */
	public ExtractedLicenseInfoSheet getExtractedLicenseInfoSheet() {
		return extractedLicenseInfoSheet;
	}

	/**
	 * @param extractedLicenseInfoSheet the extractedLicenseInfoSheet to set
	 */
	public void setExtractedLicenseInfoSheet(ExtractedLicenseInfoSheet extractedLicenseInfoSheet) {
		this.extractedLicenseInfoSheet = extractedLicenseInfoSheet;
	}

	/**
	 * @return the workbook
	 */
	public Workbook getWorkbook() {
		return workbook;
	}

	/**
	 * @return the modelStore
	 */
	public IModelStore getModelStore() {
		return modelStore;
	}

	/**
	 * @return the version
	 */
	public String getVersion() {
		return version;
	}

	/**
	 * @return the copyManager
	 */
	public ModelCopyManager getCopyManager() {
		return copyManager;
	}

	/**
	 * @param externalRefsSheet the externalRefsSheet to set
	 */
	public void setExternalRefsSheet(ExternalRefsSheet externalRefsSheet) {
		this.externalRefsSheet = externalRefsSheet;
	}

	/**
	 * Resize the height of all rows - will not exceed a maximum height
	 */
	public void resizeRow() {
		extractedLicenseInfoSheet.resizeRows();
//		originsSheet.resizeRows(); - Can't resize the origins sheet since it uses blank cells
		packageInfoSheet.resizeRows();
		perFileSheet.resizeRows();
		relationshipsSheet.resizeRows();
		annotationsSheet.resizeRows();
		if (snippetSheet != null) {
			snippetSheet.resizeRows();
		}
		if (externalRefsSheet != null) {
			externalRefsSheet.resizeRows();
		}
//		reviewersSheet.resizeRows(); - Can't resize the review sheet since it uses blank cells
	}

	/**
	 * Write the spreadsheet to the output stream
	 * @param stream
	 * @throws IOException
	 */
	public void write(OutputStream stream) throws IOException {
		this.workbook.write(stream);
	}

}
