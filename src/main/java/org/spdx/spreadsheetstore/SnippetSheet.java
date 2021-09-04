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
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.annotation.Nullable;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.ModelObject;
import org.spdx.library.model.SpdxFile;
import org.spdx.library.model.SpdxModelFactory;
import org.spdx.library.model.SpdxSnippet;
import org.spdx.library.model.SpdxSnippet.SpdxSnippetBuilder;
import org.spdx.library.model.license.AnyLicenseInfo;
import org.spdx.library.model.license.InvalidLicenseStringException;
import org.spdx.library.model.license.LicenseInfoFactory;
import org.spdx.library.model.license.SpdxNoAssertionLicense;
import org.spdx.library.model.pointer.ByteOffsetPointer;
import org.spdx.library.model.pointer.LineCharPointer;
import org.spdx.library.model.pointer.SinglePointer;
import org.spdx.library.model.pointer.StartEndPointer;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;

/**
 * @author gary
 *
 */
public class SnippetSheet extends AbstractSheet {
	static final Logger logger = LoggerFactory.getLogger(SnippetSheet.class);
	
	static final int ID_COL = 0;
	static final int NAME_COL = ID_COL + 1;
	static final int SNIPPET_FROM_FILE_ID_COL = NAME_COL + 1;
	static final int BYTE_RANGE_COL = SNIPPET_FROM_FILE_ID_COL + 1;
	static final int LINE_RANGE_COL = BYTE_RANGE_COL + 1;
	static final int CONCLUDED_LICENSE_COL = LINE_RANGE_COL + 1;
	static final int LICENSE_INFO_IN_SNIPPET_COL = CONCLUDED_LICENSE_COL + 1;
	static final int LICENSE_COMMENT_COL = LICENSE_INFO_IN_SNIPPET_COL + 1;
	static final int COPYRIGHT_COL = LICENSE_COMMENT_COL + 1;
	static final int COMMENT_COL = COPYRIGHT_COL + 1;
	static final int USER_DEFINED_COLS = COMMENT_COL + 1;
	static final int NUM_COLS = USER_DEFINED_COLS + 1;
	
	static final boolean[] REQUIRED = new boolean[] {true, false, true, true, false,
		false, false, false, false, false, false};
	static final String[] HEADER_TITLES = new String[] {"ID", "Name", "From File ID",
		"Byte Range", "Line Range", "License Concluded", "License Info in Snippet", "License Comments",
		"Snippet Copyright Text", "Comment", "User Defined Columns..."};

	static final int[] COLUMN_WIDTHS = new int[] {25, 25, 25, 40, 40, 60, 60, 60, 60, 60, 40};
	static final boolean[] LEFT_WRAP = new boolean[] {false, false, false, false, false,
		true, true, true, true, true, true};
	static final boolean[] CENTER_NOWRAP = new boolean[] {true, true, true, true, true,
		false, false, false, false, false, false};
	
	private static Pattern NUMBER_RANGE_PATTERN = Pattern.compile("(\\d+):(\\d+)");
	
	/**
	 * Hashmap of the snippet ID to SPDX snipet
	 */
	Map<String, SpdxSnippet> snippetCache = new HashMap<>();

	/**
	 * @param workbook
	 * @param snippetSheetName
	 * @param modelStore
	 * @param documentUri
	 * @param copyManager
	 */
	public SnippetSheet(Workbook workbook, String snippetSheetName, IModelStore modelStore, 
			String documentUri, ModelCopyManager copyManager) {
		super(workbook, snippetSheetName, modelStore, documentUri, copyManager);
	}

	/* (non-Javadoc)
	 * @see org.spdx.spdxspreadsheet.AbstractSheet#verify()
	 */
	@Override
	public String verify() {
		try {
			if (sheet == null) {
				return "Worksheet for SPDX Snippets does not exist";
			}
			Row firstRow = sheet.getRow(firstRowNum);
			for (int i = 0; i < NUM_COLS- 1; i++) { 	// Don't check the last (user defined) column
				Cell cell = firstRow.getCell(i+firstCellNum);
				if (cell == null || 
						cell.getStringCellValue() == null ||
						!cell.getStringCellValue().equals(HEADER_TITLES[i])) {
					return "Column "+HEADER_TITLES[i]+" missing for SPDX Snippet worksheet";
				}
			}
			// validate rows
			boolean done = false;
			int rowNum = getFirstDataRow();
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
			return "Error in verifying SPDX Snippet work sheet: "+ex.getMessage();
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
			} else {
				if (i == CONCLUDED_LICENSE_COL) {
					try {
						LicenseInfoFactory.parseSPDXLicenseString(cell.getStringCellValue(), modelStore, documentUri, copyManager);
					} catch (InvalidSPDXAnalysisException ex) {
						return "Invalid asserted license string in row "+String.valueOf(row.getRowNum()) +
								" details: "+ex.getMessage();
					}
				} else if (i == BYTE_RANGE_COL || i == LINE_RANGE_COL) {
					String range = cell.getStringCellValue();
					if (range != null && !range.isEmpty()) {
						Matcher rangeMatcher = NUMBER_RANGE_PATTERN.matcher(cell.getStringCellValue());
						if (!rangeMatcher.matches()) {
							return "Invalid range for "+HEADER_TITLES[i]+": "+cell.getStringCellValue();
						}
						int start = 0;
						int end = 0;
						try {
							start = Integer.parseInt(rangeMatcher.group(1));
							end = Integer.parseInt(rangeMatcher.group(2));
							if (start >= end) {
								return "Invalid range for "+HEADER_TITLES[i]+": "+cell.getStringCellValue() + ".  End is not greater than or equal to the end.";
							}
						} catch(Exception ex) {
							return "Invalid range for "+HEADER_TITLES[i]+": "+cell.getStringCellValue();
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
	 * @param snippet
	 * @throws SpreadsheetException 
	 */
	public void add(SpdxSnippet snippet) throws SpreadsheetException {
		Row row = addRow();
		if (snippet.getId() != null && !snippet.getId().isEmpty()) {
			row.createCell(ID_COL).setCellValue(snippet.getId());
		}
		try {
		    Optional<String> name = snippet.getName();
			if (name.isPresent()) {
				row.createCell(NAME_COL).setCellValue(name.get());
			}
			try {
				SpdxFile snippetFromFile = snippet.getSnippetFromFile();
				row.createCell(SNIPPET_FROM_FILE_ID_COL).setCellValue(snippetFromFile.getId());
			} catch (InvalidSPDXAnalysisException e) {
				logger.error("Error getting the snippetFromFile",e);
				throw new SpreadsheetException("Unable to get the Snippet from File from the Snippet: "+e.getMessage());
			}
			StartEndPointer byteRange;
			try {
				byteRange = snippet.getByteRange();
				try {
					row.createCell(BYTE_RANGE_COL).setCellValue(rangeToStr(byteRange));
				} catch (InvalidSPDXAnalysisException e) {
					logger.error("Invalid byte range",e);
					throw new SpreadsheetException("Invalid byte range: "+e.getMessage());
				}
			} catch (InvalidSPDXAnalysisException e) {
				logger.error("Error getting the byteRange",e);
				throw new SpreadsheetException("Unable to get the byte range from the Snippet: "+e.getMessage());
			}
			Optional<StartEndPointer> lineRange;
			try {
				lineRange = snippet.getLineRange();
			} catch (InvalidSPDXAnalysisException e) {
				logger.error("Error getting the lineRange",e);
				throw new SpreadsheetException("Unable to get the line range from the Snippet: "+e.getMessage());
			}
			if (lineRange.isPresent()) {
				try {
					row.createCell(LINE_RANGE_COL).setCellValue(rangeToStr(lineRange.get()));
				} catch (InvalidSPDXAnalysisException e) {
					logger.error("Invalid line range",e);
					throw new SpreadsheetException("Invalid line range: "+e.getMessage());
				}
			}
			if (snippet.getLicenseConcluded() != null) {
				row.createCell(CONCLUDED_LICENSE_COL).setCellValue(snippet.getLicenseConcluded().toString());
			}
			Collection<AnyLicenseInfo> licenseInfoFromSnippet = snippet.getLicenseInfoFromFiles();
			if (licenseInfoFromSnippet != null && licenseInfoFromSnippet.size() > 0) {
				row.createCell(LICENSE_INFO_IN_SNIPPET_COL).setCellValue(PackageInfoSheet.licensesToString(licenseInfoFromSnippet));
			}
			Optional<String> licenseComments = snippet.getLicenseComments();
			if (licenseComments.isPresent()) {
				row.createCell(LICENSE_COMMENT_COL).setCellValue(licenseComments.get());
			}
			if (snippet.getCopyrightText() != null) {
				row.createCell(COPYRIGHT_COL).setCellValue(snippet.getCopyrightText());
			}
			Optional<String> comment = snippet.getComment();
			if (comment.isPresent()) {
				row.createCell(COMMENT_COL).setCellValue(comment.get());
			}
			this.snippetCache.put(snippet.getId(), snippet);
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting SPDX Snippet information",e);
		}
	}

	/**
	 * @param byteRange
	 * @return
	 * @throws InvalidSPDXAnalysisException 
	 */
	private String rangeToStr(StartEndPointer rangePointer) throws InvalidSPDXAnalysisException {
		SinglePointer startPointer = rangePointer.getStartPointer();
		if (startPointer == null) {
			throw new InvalidSPDXAnalysisException("Missing start pointer");
		}
		SinglePointer endPointer = rangePointer.getEndPointer();
		if (endPointer == null) {
			throw new InvalidSPDXAnalysisException("Missing end pointer");
		}
		String start = null;
		if (startPointer instanceof ByteOffsetPointer) {
			start = String.valueOf(((ByteOffsetPointer)startPointer).getOffset());
		} else if (startPointer instanceof LineCharPointer) {
			start = String.valueOf(((LineCharPointer)startPointer).getLineNumber());
		} else {
			logger.error("Unknown pointer type for start pointer "+startPointer.toString());
			throw new InvalidSPDXAnalysisException("Unknown pointer type for start pointer");
		}
		String end = null;
		if (endPointer instanceof ByteOffsetPointer) {
			end = String.valueOf(((ByteOffsetPointer)endPointer).getOffset());
		} else if (endPointer instanceof LineCharPointer) {
			end = String.valueOf(((LineCharPointer)endPointer).getLineNumber());
		} else {
			logger.error("Unknown pointer type for start pointer "+startPointer.toString());
			throw new InvalidSPDXAnalysisException("Unknown pointer type for start pointer");
		}
		return start + ":" + end;
	}

	/**
	 * Get the SPDX snippet represented in the row rownum.
	 * IMPORTANT: The Snippet From File must already be in the model store.
	 * The ID from the Snippet From File can be obtained through the <code> getSnippetFileId(int rowNum)</code> method
	 * @param rowNum
	 * @return Snippet at the row rowNum or null if the row does not exist
	 * @throws SpreadsheetException 
	 */
	public @Nullable SpdxSnippet getSnippet(int rowNum) throws SpreadsheetException {
		if (sheet == null) {
			return null;
		}
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		String ver = validateRow(row);
		if (ver != null && !ver.isEmpty()) {
			throw(new SpreadsheetException(ver));
		}
		String id;
		if (Objects.nonNull(row.getCell(ID_COL)) && !row.getCell(ID_COL).getStringCellValue().trim().isEmpty()) {
			id = row.getCell(ID_COL).getStringCellValue();
		} else {
			try {
				id = modelStore.getNextId(IdType.Anonymous, documentUri);
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error getting anonymous ID");
			}
			row.getCell(ID_COL).setCellValue(id);
		}
		if (this.snippetCache.containsKey(id)) {
			return this.snippetCache.get(id);
		}
		
		String name = "";
		if (row.getCell(NAME_COL) != null) {
			name = row.getCell(NAME_COL).getStringCellValue();
		}
		
		AnyLicenseInfo concludedLicense;
		Cell concludedLicenseCell = row.getCell(CONCLUDED_LICENSE_COL);
		if (concludedLicenseCell != null && !concludedLicenseCell.getStringCellValue().trim().isEmpty()) {
			try {
				concludedLicense = LicenseInfoFactory.parseSPDXLicenseString(concludedLicenseCell.getStringCellValue(), 
						modelStore, documentUri, copyManager);
			} catch (InvalidLicenseStringException e) {
				throw new SpreadsheetException("Invalid license expression "+concludedLicenseCell.getStringCellValue(),e);
			}
		} else {
			try {
				concludedLicense = new SpdxNoAssertionLicense();
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error creating NoAssertionLicense",e);
			}
		}
		
		Collection<AnyLicenseInfo> licenseInfosFromFile = new  ArrayList<>();
		Cell seenLicenseCell = row.getCell(LICENSE_INFO_IN_SNIPPET_COL);
		if (seenLicenseCell != null && !seenLicenseCell.getStringCellValue().trim().isEmpty()) {
			for (String licenseString:seenLicenseCell.getStringCellValue().split(",")) {
				try {
					licenseInfosFromFile.add(LicenseInfoFactory.parseSPDXLicenseString(licenseString.trim(), 
							modelStore, documentUri, copyManager));
				} catch (InvalidLicenseStringException e) {
					throw new SpreadsheetException("Invalid license expression in License Infos from File: "+licenseString,e);
				}
			}
		} else {
			try {
				licenseInfosFromFile.add(new SpdxNoAssertionLicense());
			} catch (InvalidSPDXAnalysisException e) {
				throw new SpreadsheetException("Error creating NoAssertionLicense",e);
			}
		}
		String copyright;
		Cell copyrightCell = row.getCell(COPYRIGHT_COL);
		if (copyrightCell != null) {
			copyright = copyrightCell.getStringCellValue();
		} else {
			copyright = "NOASSERTION";
		}
		
		//TODO: We could create the snippetFromFile if it doesn't already exist in the model rather than failing
		String snippetFromFileId = getSnippetFileId(rowNum);
		if (Objects.isNull(snippetFromFileId) || snippetFromFileId.isEmpty()) {
			throw new SpreadsheetException("Missing required Snippet From File ID for Snippet ID "+id);
		}
		Optional<ModelObject> moFromFile;
		try {
			moFromFile = SpdxModelFactory.getModelObject(modelStore, documentUri, snippetFromFileId, copyManager);
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error getting SnippetFromFile",e);
		}
		if (!moFromFile.isPresent()) {
			throw new SpreadsheetException("Snippet from file for snippet ID "+id+" does not exist in the model.  "
					+ "It must be created before getting the snippet information from the snippet sheet.  "
					+ "Restore the file information sheet first.");
		}
		if (!(moFromFile.get() instanceof SpdxFile)) {
			throw new SpreadsheetException("Invalid type for ID "+snippetFromFileId+".  Expecting SpdxFile");
		}
		SpdxFile snippetFromFile = (SpdxFile)moFromFile.get();
		
		if (Objects.isNull(row.getCell(BYTE_RANGE_COL)) && row.getCell(BYTE_RANGE_COL).getStringCellValue().trim().isEmpty()) {
			throw new SpreadsheetException("Missing reqired byte range for Snippet ID "+id);
		}
			String range = row.getCell(BYTE_RANGE_COL).getStringCellValue();
			int start = 0;
			int end = 0;
			Matcher rangeMatcher = NUMBER_RANGE_PATTERN.matcher(range);
			if (!rangeMatcher.matches()) {
				throw new SpreadsheetException("Invalid byte range: "+range);
			}
			try {
				start = Integer.parseInt(rangeMatcher.group(1));
				end = Integer.parseInt(rangeMatcher.group(2));
			} catch(Exception ex) {
				throw new SpreadsheetException("Invalid byte range: "+range);
			}
		
		SpdxSnippetBuilder snippetBuilder = new SpdxSnippetBuilder(modelStore, documentUri, id, copyManager, 
				name, concludedLicense, licenseInfosFromFile, copyright, snippetFromFile, start, end);

		if (Objects.nonNull(row.getCell(LINE_RANGE_COL))) {
			String lineRange = row.getCell(LINE_RANGE_COL).getStringCellValue();
			if (lineRange != null && !lineRange.isEmpty()) {
				int lineStart = 0;
				int lineEnd = 0;
				Matcher lineRangeMatcher = NUMBER_RANGE_PATTERN.matcher(lineRange);
				if (!lineRangeMatcher.matches()) {
					throw new SpreadsheetException("Invalid line range: "+lineRange);
				}
				try {
					lineStart = Integer.valueOf(lineRangeMatcher.group(1));
					lineEnd = Integer.valueOf(lineRangeMatcher.group(2));
				} catch(Exception ex) {
					throw new SpreadsheetException("Invalid line range: "+lineRange);
				}
				snippetBuilder.setLineRange(lineStart, lineEnd);
			}
		}

		Cell licCommentCell = row.getCell(LICENSE_COMMENT_COL);
		if (Objects.nonNull(licCommentCell) && !licCommentCell.getStringCellValue().trim().isEmpty()) {
			snippetBuilder.setLicenseComments(licCommentCell.getStringCellValue());
		}
		
		Cell commentCell = row.getCell(COMMENT_COL);
		if (Objects.nonNull(commentCell) && !commentCell.getStringCellValue().trim().isEmpty()) {
			snippetBuilder.setComment(commentCell.getStringCellValue());
		}
		
		SpdxSnippet retval;
		try {
			retval = snippetBuilder.build();
		} catch (InvalidSPDXAnalysisException e) {
			throw new SpreadsheetException("Error building Snippet",e);
		}
		this.snippetCache.put(id, retval);
		return retval;
	}

	/**
	 * Get the SpdxFromFileSNippet for the given row
	 * @param rowNum
	 * @return
	 * @throws SpreadsheetException 
	 */
	public String getSnippetFileId(int rowNum) throws SpreadsheetException {
		Row row = sheet.getRow(rowNum);
		if (row == null) {
			return null;
		}
		String ver = validateRow(row);
		if (ver != null && !ver.isEmpty()) {
			throw(new SpreadsheetException(ver));
		}
		String id = null;
		if (row.getCell(SNIPPET_FROM_FILE_ID_COL) != null) {
			id = row.getCell(SNIPPET_FROM_FILE_ID_COL).getStringCellValue();
		}
		return id;
	}

}