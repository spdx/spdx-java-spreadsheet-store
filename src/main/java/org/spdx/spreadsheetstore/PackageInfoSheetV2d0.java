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
import java.util.List;
import java.util.Objects;
import java.util.Optional;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxVerificationHelper;
import org.spdx.library.model.Checksum;
import org.spdx.library.model.SpdxPackage;
import org.spdx.library.model.SpdxPackage.SpdxPackageBuilder;
import org.spdx.library.model.SpdxPackageVerificationCode;
import org.spdx.library.model.license.AnyLicenseInfo;
import org.spdx.library.model.license.InvalidLicenseStringException;
import org.spdx.library.model.license.LicenseInfoFactory;
import org.spdx.library.model.license.SpdxNoneLicense;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;

/**
 * Version 2.1 of the package info sheet
 * @author Gary O'Neall
 *
 */
public class PackageInfoSheetV2d0 extends PackageInfoSheet {

	int NAME_COL = 0;
	int ID_COL = NAME_COL + 1;
	int VERSION_COL = ID_COL+1;
	int MACHINE_NAME_COL = VERSION_COL+1;
	int SUPPLIER_COL = MACHINE_NAME_COL + 1;
	int ORIGINATOR_COL = SUPPLIER_COL + 1;
	int HOME_PAGE_COL = ORIGINATOR_COL + 1;
	int DOWNLOAD_URL_COL = HOME_PAGE_COL + 1;
	int PACKAGE_CHECKSUMS_COL = DOWNLOAD_URL_COL + 1;
	int FILE_VERIFICATION_VALUE_COL = PACKAGE_CHECKSUMS_COL + 1;
	int VERIFICATION_EXCLUDED_FILES_COL = FILE_VERIFICATION_VALUE_COL + 1;
	int SOURCE_INFO_COL = VERIFICATION_EXCLUDED_FILES_COL + 1;
	int DECLARED_LICENSE_COL = SOURCE_INFO_COL + 1;
	int CONCLUDED_LICENSE_COL = DECLARED_LICENSE_COL + 1;
	int LICENSE_INFO_IN_FILES_COL = CONCLUDED_LICENSE_COL + 1;
	int LICENSE_COMMENT_COL = LICENSE_INFO_IN_FILES_COL + 1;
	int DECLARED_COPYRIGHT_COL = LICENSE_COMMENT_COL + 1;
	int SHORT_DESC_COL = DECLARED_COPYRIGHT_COL + 1;
	int FULL_DESC_COL = SHORT_DESC_COL + 1;
	int USER_DEFINED_COL = FULL_DESC_COL + 1;
	int NUM_COLS = USER_DEFINED_COL;

	
	static final boolean[] REQUIRED = new boolean[] {true, true, false, false, false, false, false, true, 
		true, true, false, false, true, true, true, false, true, false, false, false};
	static final String[] HEADER_TITLES = new String[] {"Package Name", "SPDX Identifier", "Package Version", 
		"Package FileName", "Package Supplier", "Package Originator", "Home Page",
		"Package Download Location", "Package Checksum", "Package Verification Code",
		"Verification Code Excluded Files", "Source Info", "License Declared", "License Concluded", "License Info From Files", 
		"License Comments", "Package Copyright Text", "Summary", "Description", "User Defined Columns..."};
	
	static final int[] COLUMN_WIDTHS = new int[] {30, 17, 17, 30, 30, 30, 50, 50, 75, 60, 40, 30,
		40, 40, 90, 50, 50, 50, 80, 50};

	/**
	 * @param workbook
	 * @param sheetName
	 * @param version
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager 
	 */
	public PackageInfoSheetV2d0(Workbook workbook, String sheetName, String version, 
			IModelStore modelStore, String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, version, modelStore, documentUri, copyManager);
		this.version = version;
	}

	/* (non-Javadoc)
	 * @see org.spdx.rdfparser.AbstractSheet#verify()
	 */
	@Override
	public String verify() {
		try {
			if (sheet == null) {
				return "Worksheet for SPDX Package Info does not exist";
			}
			if (!SpdxSpreadsheet.verifyVersion(version)) {
				return "Unsupported version "+version;
			}
			Row firstRow = sheet.getRow(firstRowNum);
			for (int i = 0; i < NUM_COLS - 1; i++) {
				Cell cell = firstRow.getCell(i+firstCellNum);
				if (cell == null || 
						cell.getStringCellValue() == null ||
						!cell.getStringCellValue().equals(HEADER_TITLES[i])) {
					return "Column "+HEADER_TITLES[i]+" missing for SPDX Package Info worksheet";
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
			return "Unexpected error in verifying SPDX Package Info work sheet: "+ex.getMessage();
		}
	}

	private String validateRow(Row row) {
		for (int i = 0; i < NUM_COLS; i++) {
			Cell cell = row.getCell(i);
			if (cell == null) {
				if (REQUIRED[i]) {
					return "Required cell "+HEADER_TITLES[i]+" missing for row "+String.valueOf(row.getRowNum() + " in PackageInfo sheet.");
				}
			} else {
				if (i == DECLARED_LICENSE_COL || i == CONCLUDED_LICENSE_COL) {
					try {
						LicenseInfoFactory.parseSPDXLicenseString(cell.getStringCellValue(), modelStore, documentUri, copyManager);
					} catch(InvalidSPDXAnalysisException ex) {
						if (i == DECLARED_LICENSE_COL) {
							return "Invalid declared license in row "+String.valueOf(row.getRowNum())+" detail: "+ex.getMessage() + " in PackageInfo sheet.";
						} else {
							return "Invalid seen license in row "+String.valueOf(row.getRowNum())+" detail: "+ex.getMessage() + " in PackageInfo sheet.";
						}
					}
				} else if (i == LICENSE_INFO_IN_FILES_COL) {
					String[] licenses = row.getCell(LICENSE_INFO_IN_FILES_COL).getStringCellValue().split(",");
					if (licenses.length < 1) {
						return "Missing licenss information in files in PackageInfo sheet.";
					}
					for (int j = 0; j < licenses.length; j++) {
						try {
							LicenseInfoFactory.parseSPDXLicenseString(licenses[j], modelStore, documentUri, copyManager);
						} catch(InvalidSPDXAnalysisException ex) {
							return "Invalid license information in in files for license "+licenses[j]+ " row "+String.valueOf(row.getRowNum())+" detail: "+ex.getMessage() + " in PackageInfo sheet.";
						}
					}
				} else if (i == ORIGINATOR_COL) {
					Cell origCell = row.getCell(ORIGINATOR_COL);
					if (origCell != null) {
						String originator = origCell.getStringCellValue();
						if (originator != null && !originator.isEmpty()) {
							String error = SpdxVerificationHelper.verifyOriginator(originator);
							if (error != null && !error.isEmpty()) {
								return "Invalid originator in row "+String.valueOf(row.getRowNum()) + ": "+error + " in PackageInfo sheet.";
							}
						}
					}
				} else if (i == SUPPLIER_COL) {
					Cell supplierCell = row.getCell(SUPPLIER_COL);
					if (supplierCell != null) {
						String supplier = supplierCell.getStringCellValue();
						if (supplier != null && !supplier.isEmpty()) {
							String error = SpdxVerificationHelper.verifySupplier(supplier);
							if (error != null && !error.isEmpty()) {
								return "Invalid supplier in row "+String.valueOf(row.getRowNum()) + ": "+error + " in PackageInfo sheet.";
							}
						}
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
		CellStyle defaultStyle = AbstractSheet.createLeftWrapStyle(wb);
		Row row = sheet.createRow(0);
		for (int i = 0; i < HEADER_TITLES.length; i++) {
			sheet.setColumnWidth(i, COLUMN_WIDTHS[i]*256);
			sheet.setDefaultColumnStyle(i, defaultStyle);
			Cell cell = row.createCell(i);
			cell.setCellStyle(headerStyle);
			cell.setCellValue(HEADER_TITLES[i]);
		}
	}
	
	public void add(SpdxPackage pkgInfo) throws InvalidSPDXAnalysisException {
		Row row = addRow();
		Cell nameCell = row.createCell(NAME_COL);
		Optional<String> name = pkgInfo.getName();
		if (name.isPresent()) {
			nameCell.setCellValue(name.get());
		}
		Cell idCell = row.createCell(ID_COL);
		idCell.setCellValue(pkgInfo.getId());
		Cell copyrightCell = row.createCell(DECLARED_COPYRIGHT_COL);
		copyrightCell.setCellValue(pkgInfo.getCopyrightText());
		Cell DeclaredLicenseCol = row.createCell(DECLARED_LICENSE_COL);
		DeclaredLicenseCol.setCellValue(pkgInfo.getLicenseDeclared().toString());
		Cell concludedLicenseCol = row.createCell(CONCLUDED_LICENSE_COL);
		concludedLicenseCol.setCellValue(pkgInfo.getLicenseConcluded().toString());
		Cell fileChecksumCell = row.createCell(FILE_VERIFICATION_VALUE_COL);
		Optional<SpdxPackageVerificationCode> verificationCode = pkgInfo.getPackageVerificationCode();
		if (verificationCode.isPresent()) {
			fileChecksumCell.setCellValue(verificationCode.get().getValue());
			Cell verificationExcludedFilesCell = row.createCell(VERIFICATION_EXCLUDED_FILES_COL);
			StringBuilder excFilesStr = new StringBuilder();
			Collection<String> excludedFileCollection = verificationCode.get().getExcludedFileNames();
			if (excludedFileCollection.size() > 0) {
				String[] excludedFiles = excludedFileCollection.toArray(new String[excludedFileCollection.size()]);
				excFilesStr.append(excludedFiles[0]);
				for (int i = 1;i < excludedFiles.length; i++) {
					excFilesStr.append(", ");
					excFilesStr.append(excludedFiles[i]);
				}
			}
			verificationExcludedFilesCell.setCellValue(excFilesStr.toString());
		}

		Optional<String> description = pkgInfo.getDescription();
		if (description.isPresent()) {
			Cell descCell = row.createCell(FULL_DESC_COL);
			descCell.setCellValue(description.get());
		}
		Cell fileNameCell = row.createCell(MACHINE_NAME_COL);
		Optional<String> packageFileName = pkgInfo.getPackageFileName();
		if (packageFileName.isPresent()) {
			fileNameCell.setCellValue(packageFileName.get());
		}
		Cell checksumsCell = row.createCell(PACKAGE_CHECKSUMS_COL);
		Collection<Checksum> checksums = pkgInfo.getChecksums();
		checksumsCell.setCellValue(checksumsToString(checksums));
		// add the license infos in files in multiple rows
		Collection<AnyLicenseInfo> licenseInfosInFilesCollection = pkgInfo.getLicenseInfoFromFiles();
		if (licenseInfosInFilesCollection != null && licenseInfosInFilesCollection.size() > 0) {
			AnyLicenseInfo[] licenseInfosInFiles = licenseInfosInFilesCollection.toArray(new AnyLicenseInfo[licenseInfosInFilesCollection.size()]);
			StringBuilder sb = new StringBuilder(licenseInfosInFiles[0].toString());
			for (int i = 1; i < licenseInfosInFiles.length; i++) {
				sb.append(',');
				sb.append(licenseInfosInFiles[i].toString());
			}
			row.createCell(LICENSE_INFO_IN_FILES_COL).setCellValue(sb.toString());
		}
		Optional<String> licenseComment = pkgInfo.getLicenseComments();
		if (licenseComment.isPresent()) {
			row.createCell(LICENSE_COMMENT_COL).setCellValue(licenseComment.get());
		}
		Optional<String> summary = pkgInfo.getSummary();
		if (summary.isPresent()) {
			Cell shortDescCell = row.createCell(SHORT_DESC_COL);
			shortDescCell.setCellValue(summary.get());
		}
		Optional<String> sourceInfo = pkgInfo.getSourceInfo();
		if (sourceInfo.isPresent()) {
			Cell sourceInfoCell = row.createCell(SOURCE_INFO_COL);
			sourceInfoCell.setCellValue(sourceInfo.get());
		}
		Cell urlCell = row.createCell(DOWNLOAD_URL_COL);
		Optional<String> downloadLocation = pkgInfo.getDownloadLocation();
		if (downloadLocation.isPresent()) {
		    urlCell.setCellValue(downloadLocation.get());
		}
		Optional<String> version = pkgInfo.getVersionInfo();
		if (version.isPresent()) {
			Cell versionInfoCell = row.createCell(VERSION_COL);
			versionInfoCell.setCellValue(version.get());
		}
		Optional<String> originator = pkgInfo.getOriginator();
		if (originator.isPresent()) {
			Cell originatorCell = row.createCell(ORIGINATOR_COL);
			originatorCell.setCellValue(originator.get());
		}
		Optional<String> supplier = pkgInfo.getSupplier();
		if (supplier.isPresent()) {
			Cell supplierCell = row.createCell(SUPPLIER_COL);
			supplierCell.setCellValue(supplier.get());
		}
		Optional<String> homePage = pkgInfo.getHomepage();
		if (homePage.isPresent()) {
			Cell homePageCell = row.createCell(HOME_PAGE_COL);
			homePageCell.setCellValue(homePage.get());
		}
	}

	/* (non-Javadoc)
	 * @see org.spdx.spreadsheetstore.PackageInfoSheet#getPackages()
	 */
	public List<SpdxPackage> getPackages() throws SpreadsheetException {
		List<SpdxPackage> retval = new ArrayList<>();
		for (int i = 0; i < getNumDataRows(); i++) {
			retval.add(getPackage(getFirstDataRow() + i));
		}
		return retval;
	}
	
	/**
	 * @param rowNum
	 * @return SPDX package at the row rowNum, null if there is no package at that row
	 * @throws SpreadsheetException
	 */
	private SpdxPackage getPackage(int rowNum) throws SpreadsheetException {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		Cell nameCell = row.getCell(NAME_COL);
		if (nameCell == null || nameCell.getStringCellValue().isEmpty()) {
			return null;
		}
		String error = validateRow(row);
		if (error != null && !error.isEmpty()) {
			throw(new SpreadsheetException(error));
		}
		
		String declaredName = nameCell.getStringCellValue();
		String id = row.getCell(ID_COL).getStringCellValue();
		
		AnyLicenseInfo concludedLicense;
		Cell concludedLicensesCell = row.getCell(CONCLUDED_LICENSE_COL);
		if (concludedLicensesCell != null && !concludedLicensesCell.getStringCellValue().isEmpty()) {
			try {
				concludedLicense = LicenseInfoFactory.parseSPDXLicenseString(concludedLicensesCell.getStringCellValue(), modelStore, documentUri, copyManager);
			} catch (InvalidLicenseStringException e) {
				throw new SpreadsheetException("Invalid concluded license file for package "+declaredName, e);
			}
		} else {
			try {
				concludedLicense = new SpdxNoneLicense();
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Invalid license infos from file for package "+declaredName, e);
			}
		}
		
		String declaredCopyright = row.getCell(DECLARED_COPYRIGHT_COL).getStringCellValue();
		
		AnyLicenseInfo declaredLicenses;
		try {
			declaredLicenses = LicenseInfoFactory.parseSPDXLicenseString(row.getCell(DECLARED_LICENSE_COL).getStringCellValue(), modelStore, documentUri, copyManager);
		} catch (InvalidLicenseStringException e1) {
			throw new SpreadsheetException("Invalid declared license  for package "+declaredName, e1);
		}
		
		Cell checksumsCell = row.getCell(PACKAGE_CHECKSUMS_COL);
		Collection<Checksum> checksums;
		if (Objects.isNull(checksumsCell)) {
			throw new SpreadsheetException("Missing required checksum for package");
		}
		try {
			checksums = strToChecksums(checksumsCell.getStringCellValue());
		} catch (InvalidSPDXAnalysisException e) {
			throw(new SpreadsheetException("Error converting file checksums: "+e.getMessage()));
		}

		SpdxPackageBuilder retval = new SpdxPackageBuilder(modelStore, documentUri, id, copyManager, 
				declaredName, concludedLicense, declaredCopyright, declaredLicenses)
				.setChecksums(checksums);
		
		Cell machineNameCell = row.getCell(MACHINE_NAME_COL);
		if (Objects.nonNull(machineNameCell)) {
			retval.setPackageFileName(row.getCell(MACHINE_NAME_COL).getStringCellValue());
		}

		Cell sourceInfocol = row.getCell(SOURCE_INFO_COL);
		if (Objects.nonNull(sourceInfocol)) {
			retval.setSourceInfo(sourceInfocol.getStringCellValue());
		}

		Cell licenseInfoInFilesCell = row.getCell(LICENSE_INFO_IN_FILES_COL);
		if (Objects.nonNull(licenseInfoInFilesCell)) {
			String[] licenseStrings = row.getCell(LICENSE_INFO_IN_FILES_COL).getStringCellValue().split(",");
			Collection<AnyLicenseInfo> licenseInfosFromFiles = new ArrayList<>();
			for (int i = 0; i < licenseStrings.length; i++) {
				try {
					licenseInfosFromFiles.add(LicenseInfoFactory.parseSPDXLicenseString(licenseStrings[i].trim(), modelStore, documentUri, copyManager));
				} catch (InvalidLicenseStringException e) {
					throw new SpreadsheetException("Invalid license infos from file for package "+declaredName, e);
				}
			}
			retval.setLicenseInfosFromFile(licenseInfosFromFiles);
		}
		
		Cell licenseCommentCell = row.getCell(LICENSE_COMMENT_COL);
		if (Objects.nonNull(licenseCommentCell) && !licenseCommentCell.getStringCellValue().isEmpty()) {
			retval.setLicenseComments(licenseCommentCell.getStringCellValue());
		}
		
		Cell shortDescCell = row.getCell(SHORT_DESC_COL);
		if (Objects.nonNull(shortDescCell) && !shortDescCell.getStringCellValue().isEmpty()) {
			retval.setSummary(shortDescCell.getStringCellValue());
		}
		
		Cell descCell = row.getCell(FULL_DESC_COL);
		if (Objects.nonNull(descCell) && !descCell.getStringCellValue().isEmpty()) {
			retval.setDescription(descCell.getStringCellValue());
		}

		Cell downloadUrlCell = row.getCell(DOWNLOAD_URL_COL);
		if (downloadUrlCell != null) {
			retval.setDownloadLocation(downloadUrlCell.getStringCellValue());
		}

		Cell packageVerificationCell = row.getCell(FILE_VERIFICATION_VALUE_COL);
		if (Objects.nonNull(packageVerificationCell)) {
			String packageVerificationValue = packageVerificationCell.getStringCellValue();
			Collection<String> excludedFiles = new ArrayList<String>();
			
			Cell excludedFilesCell = row.getCell(VERIFICATION_EXCLUDED_FILES_COL);
			String excludedFilesStr = null;
			if (excludedFilesCell != null) {
				excludedFilesStr = excludedFilesCell.getStringCellValue();
			}
			if (excludedFilesStr != null && !excludedFilesStr.isEmpty()) {
				for (String excludedFile:excludedFilesStr.split(",")) {
					excludedFiles.add(excludedFile.trim());
				}
			}
			try {
				SpdxPackageVerificationCode verificationCode = new SpdxPackageVerificationCode(modelStore, documentUri, modelStore.getNextId(IdType.Anonymous, documentUri), copyManager, true);
				verificationCode.setValue(packageVerificationValue);
				verificationCode.getExcludedFileNames().addAll(excludedFiles);
				retval.setPackageVerificationCode(verificationCode);
			} catch(InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Invalid verification code for package "+declaredName, e);
			}
		}
		
		Cell versionInfoCell = row.getCell(VERSION_COL);
		if (Objects.nonNull(versionInfoCell)) {
			String versionInfo;
			if (versionInfoCell.getCellType()== CellType.STRING  && !versionInfoCell.getStringCellValue().isEmpty()) {
				versionInfo = versionInfoCell.getStringCellValue();
			} else if (versionInfoCell.getCellType() == CellType.NUMERIC) {
				versionInfo = Double.toString(versionInfoCell.getNumericCellValue());
			} else {
				versionInfo = "";
			}
			retval.setVersionInfo(versionInfo);
		} 

		Cell supplierCell = row.getCell(SUPPLIER_COL);
		if (Objects.nonNull(supplierCell) && !supplierCell.getStringCellValue().isEmpty()) {
			retval.setSupplier(supplierCell.getStringCellValue());
		} 

		Cell originatorCell = row.getCell(ORIGINATOR_COL);
		if (Objects.nonNull(originatorCell) && !originatorCell.getStringCellValue().isEmpty()) {
			retval.setOriginator(originatorCell.getStringCellValue());
		}

		Cell homePageCell = row.getCell(HOME_PAGE_COL);
		if (Objects.nonNull(homePageCell) && !homePageCell.getStringCellValue().isEmpty()) {
			retval.setHomepage(homePageCell.getStringCellValue());
		}
		
		try {
			return retval.build();
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error building package "+declaredName,e);
		}
	}
}