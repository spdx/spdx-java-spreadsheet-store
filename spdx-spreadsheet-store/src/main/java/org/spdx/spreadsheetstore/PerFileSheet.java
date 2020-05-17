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

import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.SpdxFile;
import org.spdx.library.model.enumerations.FileType;
import org.spdx.storage.IModelStore;

/**
 * Abstract class for PerFileSheet.  Specific version implementations are implemented
 * as subclasses.
 * 
 * @author Gary O'Neall
 */
public abstract class PerFileSheet extends AbstractSheet {
	
	protected String version;
	
	/**
	 * @param workbook
	 * @param sheetName
	 * @param version
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 */
	public PerFileSheet(Workbook workbook, String sheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, sheetName, modelStore, documentUri, copyManager);
		this.version = version;
	}
	
	/**
	 * Open a specific version of the PerFileSheet
	 * @param workbook
	 * @param perFileSheetName
	 * @param version
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 * @return
	 */
	public static PerFileSheet openVersion(Workbook workbook,
			String perFileSheetName, String version, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		if (version.compareTo(SpdxSpreadsheet.VERSION_2_1_0) <= 0) {
			return new PerFileSheetV2d0(workbook, perFileSheetName, version, modelStore, documentUri, copyManager);
		} else {
			return new PerFileSheetV2d2(workbook, perFileSheetName, version, modelStore, documentUri, copyManager);
		}
	}

	/**
	 * Add the file to the spreadsheet
	 * @param file
	 * @param pkgIds string containing the package ID's which contain this file
	 * @throws SpreadsheetException 
	 */
	public abstract void add(SpdxFile file, String pkgIds) throws SpreadsheetException;
	
	/**
	 * Get the file information for a row in the PerFileSheet
	 * @param rowNum
	 * @return
	 */
	public abstract SpdxFile getFileInfo(int rowNum) throws SpreadsheetException;

	/**	
	 * Create a blank worksheet NOTE: Replaces / deletes existing sheet by the same name
	 * @param wb
	 * @param perFileSheetName
	 */
	public static void create(Workbook wb, String perFileSheetName) {
		//NOTE: This needs to be updated the the most current version
		PerFileSheetV2d2.create(wb, perFileSheetName);
	}

	/**
	 * @param row
	 * @return
	 */
	public abstract List<String> getPackageIds(int row);
	
	/**
	 * @param fileTypes
	 * @return
	 */
	public static String fileTypesToString(Collection<FileType> fileTypeCollection) {
		if (fileTypeCollection == null || fileTypeCollection.size() == 0) {
			return "";
		}
		FileType[] fileTypes = fileTypeCollection.toArray(new FileType[fileTypeCollection.size()]);
		StringBuilder sb = new StringBuilder(fileTypes[0].toString());
		for (int i = 1;i < fileTypes.length; i++) {
			sb.append(", ");
			String fileType = fileTypes[i].toString();
			sb.append(fileType);
		}
		return sb.toString();
	}

	/**
	 * @param typeStr
	 * @return
	 * @throws InvalidSPDXAnalysisException 
	 */
	public static Collection<FileType> parseFileTypeString(String typeStr) throws InvalidSPDXAnalysisException {
		Collection<FileType> retval = new ArrayList<>();
		if (typeStr == null || typeStr.trim().isEmpty()) {
			return retval;
		}
		for (String fileTypeSub:typeStr.split(",")) {
			String fileType = fileTypeSub.trim();
			if (fileType.endsWith(",")) {
				fileType = fileType.substring(0, fileType.length()-1);
				fileType = fileType.trim();
			}
			try {
				retval.add(FileType.valueOf(fileType));
			} catch (Exception ex) {
				throw(new InvalidSPDXAnalysisException("Unrecognized file type "+fileType));
			}
		}
		return retval;
	}

}
