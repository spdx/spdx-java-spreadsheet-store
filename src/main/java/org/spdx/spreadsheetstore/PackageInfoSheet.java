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

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultStoreNotInitializedException;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.SpdxPackage;
import org.spdx.storage.IModelStore;

/**
 * Abstract PackageInfoSheet to manage cross-version implementations
 *
 * @author Gary O'Neall
 */
public abstract class PackageInfoSheet extends AbstractSheet {
	
	protected String version;

	/**
	 * Constructs a new PackageInfoSheet
	 *
	 * @param workbook    the workbook containing the sheet
	 * @param sheetName   the name of the sheet
	 * @param version     the version of the sheet
	 * @param modelStore  the model store to use
	 * @param documentUri the URI of the document
	 * @param copyManager the copy manager to use
	 */
	public PackageInfoSheet(Workbook workbook, String sheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, modelStore, documentUri, copyManager);
		this.version = version;
	}

	/**
	 * Opens an existing PackageInfoSheet
	 *
	 * @param workbook             the workbook containing the sheet
	 * @param packageInfoSheetName the name of the sheet
	 * @param version              the version of the sheet
	 * @param modelStore           the model store to use
	 * @param documentUri          the URI of the document
	 * @param copyManager          the copy manager to use
	 * @return the opened PackageInfoSheet
	 */
	public static PackageInfoSheet openVersion(Workbook workbook,
			String packageInfoSheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		if (version.compareTo(SpdxSpreadsheet.VERSION_2_0_0) <= 0) {
			return new PackageInfoSheetV2d0(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
		} else if (version.compareTo(SpdxSpreadsheet.VERSION_2_1_0) <= 0) {
			return new PackageInfoSheetV2d1(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
		} else if (version.compareTo(SpdxSpreadsheet.VERSION_2_2_0) <= 0) {
			return new PackageInfoSheetV2d2(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
		} else {
			return new PackageInfoSheetV2d3(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
		}
	}

	/**
	 * Creates a new PackageInfoSheet in the provided workbook with the specified
	 * sheet name
	 *
	 * @param wb        the workbook where the sheet will be created
	 * @param sheetName the name of the sheet to be created
	 */
	public static void create(Workbook wb, String sheetName) {
		PackageInfoSheetV2d3.create(wb, sheetName);
	}

	/**
	 * Retrieves a list of SPDX packages from the sheet
	 *
	 * @return a list of SPDX packages
	 * @throws SpreadsheetException                if there is an error reading the
	 *                                             spreadsheet
	 * @throws DefaultStoreNotInitializedException if the model store is not
	 *                                             initialized
	 */
	public abstract List<SpdxPackage> getPackages() throws SpreadsheetException, DefaultStoreNotInitializedException;

	/**
	 * Adds a new SPDX package to the sheet
	 *
	 * @param pkgInfo the SPDX package to add
	 * @throws InvalidSPDXAnalysisException if there is an error with the SPDX
	 *                                      analysis
	 */
	public abstract void add(SpdxPackage pkgInfo) throws InvalidSPDXAnalysisException;
}
