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

import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultStoreNotInitialized;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.library.LicenseInfoFactory;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.Checksum;
import org.spdx.library.model.v2.SpdxConstantsCompatV2;
import org.spdx.library.model.v2.SpdxFile;
import org.spdx.library.model.v2.SpdxFile.SpdxFileBuilder;
import org.spdx.library.model.v2.SpdxPackage;
import org.spdx.library.model.v2.SpdxPackage.SpdxPackageBuilder;
import org.spdx.library.model.v2.enumerations.ChecksumAlgorithm;
import org.spdx.library.model.v2.enumerations.RelationshipType;
import org.spdx.library.model.v2.license.AnyLicenseInfo;
import org.spdx.library.model.v2.license.InvalidLicenseStringException;
import org.spdx.library.model.v2.license.SpdxNoAssertionLicense;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;

/**
 * Per file sheet voer version 2.2 of the SPDX spec
 * 
 * @author Gary O'Neall
 *
 */
public class PerFileSheetV2d3 extends PerFileSheet {
	
	static final int NUM_COLS = 18;
	static final int FILE_NAME_COL = 0;
	static final int ID_COL = FILE_NAME_COL + 1;
	static final int PACKAGE_ID_COL = ID_COL + 1;
	static final int FILE_TYPE_COL = PACKAGE_ID_COL + 1;
	static final int CHECKSUMS_COL = FILE_TYPE_COL + 1;
	static final int CONCLUDED_LIC_COL = CHECKSUMS_COL + 1;
	static final int LIC_INFO_IN_FILE_COL = CONCLUDED_LIC_COL + 1;
	static final int LIC_COMMENTS_COL = LIC_INFO_IN_FILE_COL + 1;
	static final int SEEN_COPYRIGHT_COL = LIC_COMMENTS_COL + 1;
	static final int NOTICE_TEXT_COL = SEEN_COPYRIGHT_COL + 1;
	static final int ARTIFACT_OF_PROJECT_COL = NOTICE_TEXT_COL + 1;
	static final int ARTIFACT_OF_HOMEPAGE_COL = ARTIFACT_OF_PROJECT_COL + 1;
	static final int ARTIFACT_OF_PROJECT_URL_COL = ARTIFACT_OF_HOMEPAGE_COL + 1;
	static final int CONTRIBUTORS_COL = ARTIFACT_OF_PROJECT_URL_COL + 1;
	static final int COMMENT_COL = CONTRIBUTORS_COL + 1;
	static final int FILE_DEPENDENCIES_COL = COMMENT_COL + 1;
	static final int ATTRIBUTION_COL = FILE_DEPENDENCIES_COL + 1;
	static final int USER_DEFINED_COL = ATTRIBUTION_COL + 1;
	
	static final boolean[] REQUIRED = new boolean[] {true, true, false, false, true, false, 
		false, false, false, false, false, false, false, false, false, false, false, false};
	static final String[] HEADER_TITLES = new String[] {"File Name", "SPDX Identifier",
		"Package Identifier", "File Type(s)",
		"File Checksum(s)", "License Concluded", "License Info in File", "License Comments",
		"File Copyright Text", "Notice Text", "Artifact of Project", "Artifact of Homepage", 
		"Artifact of URL", "Contributors", "File Comment", "File Dependencies", 
		"Attribution Text", "User Defined Columns..."};
	static final int[] COLUMN_WIDTHS = new int[] {60, 25, 25, 30, 85, 50, 50, 60,
		70, 70, 35, 60, 60, 60, 60, 60, 60, 60};
	static final boolean[] LEFT_WRAP = new boolean[] {true, false, false, true, true, 
		true, true, true, true, true, true, true, true, true, true, true, true, true};
	static final boolean[] CENTER_NOWRAP = new boolean[] {false, true, true, false, false, 
		false, false, false, false, false, false, false, false, false, false, false, false, false};
	
	/**
	 * Hashmap of the file name to SPDX file
	 */
	Map<String, SpdxFile> fileCache = new HashMap<>();
	
	PerFileSheetV2d3(Workbook workbook, String sheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, version, modelStore, documentUri, copyManager);
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
	
	/* (non-Javadoc)
	 * @see org.spdx.spreadsheetstore.PerFileSheet#add(org.spdx.library.model.SpdxFile, java.lang.String)
	 */
	@Override
	public void add(SpdxFile fileInfo, String pkgId) throws SpreadsheetException {
		Row row = addRow();
		if (fileInfo.getId() != null && !fileInfo.getId().isEmpty()) {
			row.createCell(ID_COL).setCellValue(fileInfo.getId());
		}
		if (pkgId != null && !pkgId.isEmpty()) {
			row.createCell(PACKAGE_ID_COL).setCellValue(pkgId);
		}
		// Note: this version of the library does not support artifactOf
		try {
			if (fileInfo.getLicenseConcluded() != null) {
				row.createCell(CONCLUDED_LIC_COL).setCellValue(fileInfo.getLicenseConcluded().toString());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting concluded license from file ID "+fileInfo.getId(),e);
		}
		try {
		    Optional<String> name = fileInfo.getName();
			if (name.isPresent()) {
				row.createCell(FILE_NAME_COL).setCellValue(name.get());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting file name from file ID "+fileInfo.getId(),e);
		}
		if (fileInfo.getChecksums().size() > 0) {
			try {
				row.createCell(CHECKSUMS_COL).setCellValue(checksumsToString(fileInfo.getChecksums()));
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error getting checksums from file ID "+fileInfo.getId(),e);
			}
		}
		try {
			row.createCell(FILE_TYPE_COL).setCellValue(
					fileTypesToString(fileInfo.getFileTypes()));
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting file types from file ID "+fileInfo.getId(),e);
		}
		try {
		    Optional<String> licenseComments = fileInfo.getLicenseComments();
			if (licenseComments.isPresent() && !licenseComments.get().isEmpty()) {
				row.createCell(LIC_COMMENTS_COL).setCellValue(licenseComments.get());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting license comment from file ID "+fileInfo.getId(),e);
		}
		try {
			if (fileInfo.getCopyrightText() != null && !fileInfo.getCopyrightText().isEmpty()) {
				row.createCell(SEEN_COPYRIGHT_COL).setCellValue(fileInfo.getCopyrightText());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting copyright text from file ID "+fileInfo.getId(),e);
		}
		try {
			if (fileInfo.getLicenseInfoFromFiles() != null && fileInfo.getLicenseInfoFromFiles().size() > 0) {
				row.createCell(LIC_INFO_IN_FILE_COL).setCellValue(licensesToString(fileInfo.getLicenseInfoFromFiles()));
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting license info from files from file ID "+fileInfo.getId(),e);
		}
		try {
		    Optional<String> comment = fileInfo.getComment();
			if (comment.isPresent() && !comment.get().isEmpty()) {
				row.createCell(COMMENT_COL).setCellValue(comment.get());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting comment file ID "+fileInfo.getId(),e);
		}
		if (fileInfo.getFileContributors() != null && fileInfo.getFileContributors().size() > 0) {
			row.createCell(CONTRIBUTORS_COL).setCellValue(stringsToCsv(fileInfo.getFileContributors()));	
		}
		try {
			if (fileInfo.getAttributionText() != null && fileInfo.getAttributionText().size() > 0) {
				row.createCell(ATTRIBUTION_COL).setCellValue(stringsToCsv(fileInfo.getAttributionText()));
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting attribution text from file ID "+fileInfo.getId(),e);
		}
		// Note: this version of the model does not support package dependencies
		try {
		    Optional<String> noticeText = fileInfo.getNoticeText();
			if (noticeText.isPresent() && !noticeText.get().isEmpty()) {
				row.createCell(NOTICE_TEXT_COL).setCellValue(noticeText.get());
			}
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting file notice text from file ID "+fileInfo.getId(),e);
		}
	}

	/* (non-Javadoc)
	 * @see org.spdx.spreadsheetstore.PerFileSheet#getFileInfo(int)
	 */
	@Override
	public SpdxFile getFileInfo(int rowNum) throws SpreadsheetException, DefaultStoreNotInitialized {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		String ver = validateRow(row);
		if (ver != null && !ver.isEmpty()) {
			throw(new SpreadsheetException(ver));
		}
		String name = row.getCell(FILE_NAME_COL).getStringCellValue();
		
		if (this.fileCache.containsKey(name)) {
			return this.fileCache.get(name);
		}
		
		Cell idCell = row.getCell(ID_COL);
		String id;
		if (Objects.nonNull(idCell) && idCell.getStringCellValue() != null && !idCell.getStringCellValue().isEmpty()) {
			id = idCell.getStringCellValue().trim();
		} else {
			try {
				id = modelStore.getNextId(IdType.Anonymous);
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error getting element ID for file "+name, e);
			}
		}
		
		Cell checksumsCell = row.getCell(CHECKSUMS_COL);
		Collection<Checksum> checksums = new ArrayList<>();
		Checksum sha1 = null;
		if (Objects.nonNull(checksumsCell)) {
			try {
				checksums = strToChecksums(checksumsCell.getStringCellValue());
			} catch (InvalidSPDXAnalysisException e) {
				throw(new SpreadsheetException("Error converting file checksums: "+e.getMessage(), e));
			}
		}
		for (Checksum checksum:checksums) {
			try {
				if (ChecksumAlgorithm.SHA1.equals(checksum.getAlgorithm())) {
					if (Objects.isNull(sha1)) {
						sha1 = checksum;
					} else {
						throw new SpreadsheetException("Duplicate SHA1 for file "+name);
					}
				}
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error getting checksum for file "+name, e);
			}
		}
		if (Objects.isNull(sha1)) {
			throw new SpreadsheetException("Missing SHA1 for file "+name);
		}
		
		AnyLicenseInfo concludedLicense = null;
		Cell concludedLicenseCell = row.getCell(CONCLUDED_LIC_COL);
		if (Objects.nonNull(concludedLicenseCell) && !concludedLicenseCell.getStringCellValue().isEmpty()) {
			try {
				concludedLicense = LicenseInfoFactory.parseSPDXLicenseStringCompatV2(concludedLicenseCell.getStringCellValue(), 
						modelStore, documentUri, copyManager);
			} catch (InvalidLicenseStringException e) {
				throw new SpreadsheetException("Error getting concluded license for file "+name, e);
			}
		}
		Collection<AnyLicenseInfo> licenseInfosFromFile = new ArrayList<>();
		Cell licenseInfoFromFileCell = row.getCell(LIC_INFO_IN_FILE_COL);
		if (Objects.nonNull(licenseInfoFromFileCell) && !licenseInfoFromFileCell.getStringCellValue().isEmpty()) {
			String[] licenseStrings = licenseInfoFromFileCell.getStringCellValue().split(",");
			for (int i = 0; i < licenseStrings.length; i++) {
				try {
					licenseInfosFromFile.add(LicenseInfoFactory.parseSPDXLicenseStringCompatV2(licenseStrings[i].trim(),
							modelStore, documentUri, copyManager));
				} catch (InvalidLicenseStringException e) {
					throw new SpreadsheetException("Error getting license infos from file for file "+name, e);
				}
			}
		}
		
		String copyrightText = null;
		Cell copyrightCell = row.getCell(SEEN_COPYRIGHT_COL);
		if (Objects.nonNull(copyrightCell)) {
			copyrightText = copyrightCell.getStringCellValue();
		}
		
		SpdxFileBuilder fileBuilder = new SpdxFileBuilder(modelStore, documentUri, id, copyManager, 
				name, concludedLicense, licenseInfosFromFile, copyrightText, sha1);
		for (Checksum checksum:checksums) {
			try {
				if (!ChecksumAlgorithm.SHA1.equals(checksum.getAlgorithm())) {
					fileBuilder.addChecksum(checksum);
				}
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error getting checksum value",e);
			}
		}
		
		String typeStr = row.getCell(FILE_TYPE_COL).getStringCellValue();
		if (Objects.nonNull(typeStr)) {
			try {
				fileBuilder.setFileTypes(parseFileTypeString(typeStr));
			} catch (InvalidSPDXAnalysisException e1) {
				throw(new SpreadsheetException("Error converting file types: "+e1.getMessage()));
			}
		}

		Cell licCommentCell = row.getCell(LIC_COMMENTS_COL);
		if (Objects.nonNull(licCommentCell)) {
			fileBuilder.setLicenseComments(licCommentCell.getStringCellValue());
		}
		
		Cell contributorCell = row.getCell(CONTRIBUTORS_COL);
		if (Objects.nonNull(contributorCell) && !contributorCell.getStringCellValue().trim().isEmpty()) {
			fileBuilder.setFileContributors(csvToStrings(contributorCell.getStringCellValue().trim()));
		}
		
		Cell attributionCell = row.getCell(ATTRIBUTION_COL);
		if (Objects.nonNull(attributionCell) && !attributionCell.getStringCellValue().trim().isEmpty()) {
			fileBuilder.setAttributionText(csvToStrings(attributionCell.getStringCellValue().trim()));
		}

		Cell noticeCell = row.getCell(NOTICE_TEXT_COL);
		if (Objects.nonNull(noticeCell) && !noticeCell.getStringCellValue().trim().isEmpty()) {
			fileBuilder.setNoticeText(noticeCell.getStringCellValue().trim());
		}
		
		Cell commentCell = row.getCell(COMMENT_COL);
		if (Objects.nonNull(commentCell) && !commentCell.getStringCellValue().trim().isEmpty()) {
			fileBuilder.setComment(commentCell.getStringCellValue().trim());
		}

		SpdxFile retval;
		try {
			retval = fileBuilder.build();
		} catch (InvalidSPDXAnalysisException e) {
			throw(new SpreadsheetException("Error creating new SPDX file: "+e.getMessage()));
		}
		
		//artifactOf - We'll convert these to relationships
		Cell artifactOfNameCell = row.getCell(ARTIFACT_OF_PROJECT_COL);
		if (Objects.nonNull(artifactOfNameCell) && !artifactOfNameCell.getStringCellValue().isEmpty()) {
			List<String> projectNames = csvToStrings(artifactOfNameCell.getStringCellValue());
			Cell artifactOfHomePageCell = row.getCell(ARTIFACT_OF_HOMEPAGE_COL);
			List<String> projectHomePages;
			if (Objects.nonNull(artifactOfHomePageCell) && !artifactOfHomePageCell.getStringCellValue().isEmpty()) {
				projectHomePages = csvToStrings(artifactOfHomePageCell.getStringCellValue());
			} else {
				projectHomePages = new ArrayList<String>();
			}
			AnyLicenseInfo noAssertion;
			try {
				noAssertion = new SpdxNoAssertionLicense();
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error creating no assertion license for DOAP project for file "+name, e);
			}
			for (int i = 0; i < projectNames.size(); i++) {
				SpdxPackageBuilder pkgBuilder = new SpdxPackageBuilder(modelStore, documentUri, 
						SpdxConstantsCompatV2.SPDX_ELEMENT_REF_PRENUM + "FromDoap-"+Integer.toString(i), 
						copyManager, projectNames.get(i), noAssertion, SpdxConstantsCompatV2.NOASSERTION_VALUE,
						noAssertion)
						.setFilesAnalyzed(false);
				if (projectHomePages.size() > i) {
					pkgBuilder.setHomepage(projectHomePages.get(i));
				}
				pkgBuilder.setComment("This package was converted from a DOAP Project by the same name");
				SpdxPackage doapPackage;
				try {
					doapPackage = pkgBuilder.build();
				} catch (InvalidSPDXAnalysisException e) {
					throw new SpreadsheetException("Error creating package for DOAP project for file "+name, e);
				}
				try {
					retval.addRelationship(retval.createRelationship(doapPackage, RelationshipType.GENERATED_FROM, 
							"This relationship replaces an ArtifactOf"));
				} catch (InvalidSPDXAnalysisException e) {
					throw new SpreadsheetException("Error creating relationships for DOAP project for file "+name, e);
				}
			}
		}
		
		// File dependencies - we'll convert these to relationships
		Cell fileDependencyCells = row.getCell(FILE_DEPENDENCIES_COL);
		if (Objects.nonNull(fileDependencyCells) && !fileDependencyCells.getStringCellValue().isEmpty()) {
			for (String dependencyName:csvToStrings(fileDependencyCells.getStringCellValue())) {
				SpdxFile dependency = findFileByName(dependencyName.trim());
				try {
					retval.addRelationship(retval.createRelationship(dependency, 
							RelationshipType.DEPENDS_ON, "This relationship replaced a file dependency property value"));
				} catch (InvalidSPDXAnalysisException e) {
					throw new SpreadsheetException("Error creating relationship for file dependency for file "+name, e);
				}
			}
		}
		this.fileCache.put(name, retval);
		return retval;
	}
	
	/**
	 * Finds an SPDX file by name by searching through the rows for a matching file name
	 * @param fileName
	 * @return
	 * @throws SpreadsheetException 
	 * @throws DefaultStoreNotInitialized 
	 */
	public SpdxFile findFileByName(String fileName) throws SpreadsheetException, DefaultStoreNotInitialized {
		if (this.fileCache.containsKey(fileName)) {
			return this.fileCache.get(fileName);
		}
		for (int i = this.firstRowNum; i < this.lastRowNum+1; i++) {
			Cell fileNameCell = sheet.getRow(i).getCell(FILE_NAME_COL);
			if (fileNameCell.getStringCellValue().trim().equals(fileName)) {
				return getFileInfo(i);	//note: this will add the file to the cache
			}
		}
		throw(new SpreadsheetException("Could not find dependant file in the spreadsheet: "+fileName));
	}

	/* (non-Javadoc)
	 * @see org.spdx.spreadsheetstore.PerFileSheet#getPackageIds(int)
	 */
	@Override
	public List<String> getPackageIds(int row) {
		List<String> retval = new ArrayList<>();
		Cell pkgIdCell = sheet.getRow(row).getCell(PACKAGE_ID_COL);
		if (pkgIdCell == null || pkgIdCell.getStringCellValue() == null ||
				pkgIdCell.getStringCellValue().isEmpty()) {
			return retval;
		}
		for (String pkgId:pkgIdCell.getStringCellValue().split(",")) {
			retval.add(pkgId.trim());
		}
		return retval;
	}

	/* (non-Javadoc)
	 * @see org.spdx.spreadsheetstore.AbstractSheet#verify()
	 */
	@Override
	public String verify() {
		try {
			if (sheet == null) {
				return "Worksheet for SPDX File does not exist";
			}
			Row firstRow = sheet.getRow(firstRowNum);
			for (int i = 0; i < NUM_COLS-1; i++) { 	// Don't check the last (user defined) column
				Cell cell = firstRow.getCell(i+firstCellNum);
				if (cell == null || 
						cell.getStringCellValue() == null ||
						!cell.getStringCellValue().equals(HEADER_TITLES[i])) {
					return "Column "+HEADER_TITLES[i]+" missing for SPDX File worksheet";
				}
			}
			// validate rows
			boolean done = false;
			int rowNum = firstRowNum + 1;
			while (!done) {
				Row row = sheet.getRow(rowNum);
                if (row == null || row.getCell(firstCellNum) == null || 
                        row.getCell(firstCellNum).getCellType() == CellType.BLANK ||
                        (row.getCell(firstCellNum).getCellType() == CellType.STRING && row.getCell(firstCellNum).getStringCellValue().trim().isEmpty())) {
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
			return "Error in verifying SPDX File work sheet: "+ex.getMessage();
		}
	}

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
	
	@Override
	protected String getFileId(int rowNum) {
		Cell idCell = sheet.getRow(rowNum).getCell(ID_COL);
		if (idCell == null || idCell.getStringCellValue().isEmpty()) {
			return "";
		} else {
			return idCell.getStringCellValue();
		}
	}

}
