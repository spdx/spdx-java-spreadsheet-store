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

import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.SpdxPackage;
import org.spdx.storage.IModelStore;

/**
 * Abstract PackageInfoSheet to manage cross-version implementations
 * @author Gary O'Neall
 *
 */
public abstract class PackageInfoSheet extends AbstractSheet {
	
	protected String version;

	/**
	 * @param workbook
	 * @param sheetName
	 * @param version
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 */
	public PackageInfoSheet(Workbook workbook, String sheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, modelStore, documentUri, copyManager);
		this.version = version;
	}
	
	public static void create(Workbook wb, String sheetName) {
		PackageInfoSheetV2d2.create(wb, sheetName);
	}

	/**
	 * Opens an existing PackageInfoSheet
	 * @param workbook
	 * @param packageInfoSheetName
	 * @param version
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 * @return
	 */
	public static PackageInfoSheet openVersion(Workbook workbook,
			String packageInfoSheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		if (version.compareTo(SpdxSpreadsheet.VERSION_2_0_0) <= 0) {
			return new PackageInfoSheetV2d0(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
		} else if (version.compareTo(SpdxSpreadsheet.VERSION_2_1_0) <= 0) {
			return new PackageInfoSheetV2d1(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
		} else {
			return new PackageInfoSheetV2d2(workbook, packageInfoSheetName, version, modelStore, documentUri, copyManager);
		}
	}

	/**
	 * @return
	 * @throws SpreadsheetException
	 */
	public abstract List<SpdxPackage> getPackages() throws SpreadsheetException;
	
	/**
	 * @param pkgInfo
	 * @throws InvalidSPDXAnalysisException
	 */
	public abstract void add(SpdxPackage pkgInfo) throws InvalidSPDXAnalysisException;
}