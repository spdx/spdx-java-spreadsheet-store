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

import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.ModelCopyManager;
import org.spdx.storage.IModelStore;

/**
 * Abstract class for extracted license info sheet.  Specific versions are implemented as subclasses.
 * @author Gary O'Neall
 *
 */
public abstract class ExtractedLicenseInfoSheet extends AbstractSheet {
	
	protected String version;
	
	
	/**
	 * @param workbook
	 * @param sheetName
	 * @param version spreadsheet version
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 */

	/**
	 * Create an {@link ExtractedLicenseInfoSheet} instance
	 * 
	 * @param workbook Workbook object where the sheet resides.
	 * @param sheetName Sheet name within the workbook.
	 * @param version Spreadsheet format version.
	 * @param modelStore Model store for SPDX objects.
	 * @param documentUri URI of the SPDX document.
	 * @param copyManager Copy manager for SPDX object handling.
	 */
	public ExtractedLicenseInfoSheet(Workbook workbook, String sheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, modelStore, documentUri, copyManager);
		this.version = version;
	}


	/**
	 * Opens an existing extracted license info sheet for a specific version
	 *
	 * @param workbook Workbook containing the sheet.
	 * @param packageInfoSheetName Name of the sheet containing package information.
	 * @param version Spreadsheet format version.
	 * @param modelStore Model store for SPDX objects.
	 * @param documentUri URI of the SPDX document.
	 * @param copyManager Copy manager for SPDX object handling.
	 * @return An instance of {@link ExtractedLicenseInfoSheet}.
	 */
	public static ExtractedLicenseInfoSheet openVersion(Workbook workbook,
			String packageInfoSheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		return new ExtractedLicenseInfoSheetV1d1(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
	}
	
	/**
	 * Create a blank worksheet
	 * 
	 * NOTE: Replaces / deletes existing sheet by the same name
	 * 
	 * @param wb Workbook where the sheet will be created.
	 * @param nonStandardLicenseSheetName Name of the sheet to be created.
	 */
	public static void create(Workbook wb, String nonStandardLicenseSheetName) {
		//NOTE: This needs to be updated to the current version
		ExtractedLicenseInfoSheetV1d1.create(wb, nonStandardLicenseSheetName);
	}

	
	/**
	 * Retrieve the license identifier for a specific row in the spreadsheet
	 *
	 * @param rowNum Row number in the sheet.
	 * @return License identifier.
	 */
	public abstract String getIdentifier(int rowNum);
	
	/**
	 * Retrieve the extracted text for a specific row in the spreadsheet
	 *
	 * @param rowNum Row number in the sheet.
	 * @return Extracted license text.
	 */
	public abstract String getExtractedText(int rowNum);
	
	/**
	 * Add a new row to the NonStandardLicenses sheet
	 *
	 * @param identifier License ID
	 * @param extractedText Extracted license text
	 * @param licenseName optional license name
	 * @param crossRefUrls optional cross reference URL's
	 * @param comment optional comment
	 */
	public abstract void add(String identifier, String extractedText, String licenseName,
			Collection<String> crossRefUrls, String comment);
	
	/**
	 * Retrieves the license name for a specific row.
	 *
	 * @param rowNum Row number in the sheet.
	 * @return License name.
	 */
	public abstract String getLicenseName(int rowNum);

	/**
	 * Retrieves the cross-reference URLs for a specific row.
	 *
	 * @param rowNum Row number in the sheet.
	 * @return A collection of cross-reference URLs.
	 */
	public abstract Collection<String> getCrossRefUrls(int rowNum);

	/**
	 * Retrieves the comment for a specific row.
	 *
	 * @param rowNum Row number in the sheet.
	 * @return Comment as a string.
	 */
	public abstract String getComment(int rowNum);

}
