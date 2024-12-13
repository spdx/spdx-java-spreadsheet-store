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

import java.awt.font.FontRenderContext;
import java.awt.font.TextAttribute;
import java.io.IOException;
import java.io.StringReader;
import java.io.StringWriter;
import java.text.AttributedString;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.annotation.Nullable;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.Checksum;
import org.spdx.library.model.v2.enumerations.ChecksumAlgorithm;
import org.spdx.library.model.v2.license.AnyLicenseInfo;
import org.spdx.storage.IModelStore;

import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.CSVWriter;
import com.opencsv.exceptions.CsvValidationException;

/**
 * Abstract class representing a workbook sheet used in storing structured data
 * @author Gary O'Neall
 *
 */
public abstract class AbstractSheet {
	
    static final int MAX_CHARACTERS_PER_CELL = 32767;
	static final char CSV_SEPARATOR_CHAR = ',';
	static final char CSV_QUOTING_CHAR = '"';
	static final char CSV_ESCAPE_CHAR = '\\';
	static final String CSV_LINE_END = CSVWriter.DEFAULT_LINE_END;
	private static final CSVParser parser = new CSVParserBuilder()
			.withEscapeChar(CSV_ESCAPE_CHAR)
			.withQuoteChar(CSV_QUOTING_CHAR)
			.withSeparator(CSV_SEPARATOR_CHAR)
			.build();
	
	public static Pattern CHECKSUM_PATTERN = Pattern.compile("(\\S+):\\s+(\\S+)");
	
	// Default style for cells
	static final String FONT_NAME = "Arial";
	protected static final short FONT_SIZE = (short)10*20;
	static final String CHECKBOX_FONT_NAME = "Wingdings 2";
	static final String CHECKBOX = "P";
	private static final short MAX_ROW_LINES = 10;
	protected CellStyle checkboxStyle;
	protected CellStyle dateStyle;
	protected CellStyle greenWrapped;
	protected CellStyle redWrapped;
	protected CellStyle yellowWrapped;
	
	protected Workbook workbook;
	protected Sheet sheet;
	protected int lastRowNum;
	protected int firstCellNum;
	protected int firstRowNum;
	
	protected IModelStore modelStore;
	protected String documentUri;
	protected ModelCopyManager copyManager;

	/**
	 * @param workbook Workbook where the sheet lives
	 * @param sheetName Name of the sheet
	 * @param modelStore Model store for creating typed values
	 * @param documentUri URI for the document if known
	 * @param copyManager 
	 */
	public AbstractSheet(Workbook workbook, String sheetName, IModelStore modelStore, 
			@Nullable String documentUri, ModelCopyManager copyManager) {
		Objects.requireNonNull(workbook, "Missing required workbook");
		Objects.requireNonNull(sheetName, "Missing required sheetName");
		Objects.requireNonNull(modelStore, "Missing required modelStore");
		Objects.requireNonNull(copyManager, "Missing required copyManager");
		this.modelStore = modelStore;
		this.documentUri = documentUri;
		this.workbook = workbook;
		this.copyManager = copyManager;
		sheet = workbook.getSheet(sheetName);
		if (sheet != null) {
			firstRowNum = sheet.getFirstRowNum();
			Row firstRow = sheet.getRow(firstRowNum);
			if (firstRow == null) {
				firstCellNum = 1;
			} else {
				firstCellNum = firstRow.getFirstCellNum();
			}
			findLastRow();
		} else {
			firstRowNum = 0;
			lastRowNum = 0;
			firstCellNum = 0;
		}
		createStyles(workbook);
	}
	
	/**
	 * create the styles in the workbook
	 */
	private void createStyles(Workbook wb) {
		// create the styles
		this.checkboxStyle = wb.createCellStyle();
		this.checkboxStyle.setAlignment(HorizontalAlignment.CENTER);
		this.checkboxStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		this.checkboxStyle.setBorderBottom(BorderStyle.THIN);
		this.checkboxStyle.setBorderLeft(BorderStyle.THIN);
		this.checkboxStyle.setBorderRight(BorderStyle.THIN);
		this.checkboxStyle.setBorderTop(BorderStyle.THIN);
		Font checkboxFont = wb.createFont();
		checkboxFont.setFontHeight(FONT_SIZE);
		checkboxFont.setFontName(CHECKBOX_FONT_NAME);
		this.checkboxStyle.setFont(checkboxFont);
		
		this.dateStyle = wb.createCellStyle();
		DataFormat df = wb.createDataFormat();
		this.dateStyle.setDataFormat(df.getFormat("m/d/yy h:mm"));
		
		this.greenWrapped = createLeftWrapStyle(wb);
		this.greenWrapped.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		this.greenWrapped.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		this.greenWrapped.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		this.yellowWrapped = createLeftWrapStyle(wb);
		this.yellowWrapped.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
		this.yellowWrapped.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		this.redWrapped = createLeftWrapStyle(wb);
		this.redWrapped.setFillForegroundColor(IndexedColors.RED.getIndex());
		this.redWrapped.setFillPattern(FillPatternType.SOLID_FOREGROUND);
	}
	
	/**
	 * 
	 */
	private void findLastRow() {
		boolean done = false;
		lastRowNum = firstRowNum + 1;
		try {
			while (!done) {
				Row row = sheet.getRow(lastRowNum);
				if (row == null || row.getCell(firstCellNum) == null || 
						row.getCell(firstCellNum).getStringCellValue() == null || 
						row.getCell(firstCellNum).getCellType() == CellType.BLANK ||
						row.getCell(firstCellNum).getStringCellValue().trim().isEmpty()) {
					lastRowNum--;
					done = true;
				} else {
					lastRowNum++;
				}
			}
		}
		catch (Exception ex) {
			// we just stop - stop counting rows at the first invalid row
		}
	}
	
	/**
	 * Add a new row to the end of the sheet
	 * @return new row
	 */
	protected Row addRow() {
		lastRowNum++;
		Row row = sheet.createRow(lastRowNum);
		return row;
	}
	
	/**
	 * Clears all data from the worksheet
	 */
	public void clear() {
		for (int i = lastRowNum; i > firstRowNum; i--) {
			Row row = sheet.getRow(i);
			sheet.removeRow(row);
		}
		lastRowNum = firstRowNum;
	}	
	
	public int getFirstDataRow() {
		return this.firstRowNum + 1;
	}
	
	public int getNumDataRows() {
		return this.lastRowNum - (this.firstRowNum);
	}
	
	public Sheet getSheet() {
		return this.sheet;
	}
	
	public abstract String verify();

	public static CellStyle createHeaderStyle(Workbook wb) {
		CellStyle headerStyle = wb.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font headerFont = wb.createFont();
		headerFont.setFontName("Arial");
		headerFont.setFontHeight(FONT_SIZE);
		headerFont.setBold(true);
		headerStyle.setFont(headerFont);
		headerStyle.setAlignment(HorizontalAlignment.CENTER);
		headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		headerStyle.setWrapText(true);
		return headerStyle;
	}
	
	public static CellStyle createLeftWrapStyle(Workbook wb) {
		CellStyle wrapStyle = wb.createCellStyle();
		wrapStyle.setWrapText(true);
		wrapStyle.setAlignment(HorizontalAlignment.LEFT);
		wrapStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return wrapStyle;
	}
	
	public static CellStyle createCenterStyle(Workbook wb) {
		CellStyle centerStyle = wb.createCellStyle();
		centerStyle.setWrapText(false);
		centerStyle.setAlignment(HorizontalAlignment.CENTER);
		return centerStyle;
	}
	
	/**
	 * resize the rows for a best fit.  Will not exceed maximum row height.
	 */
	public void resizeRows() {
		// header row
		// data rows
		int lastRow = this.getNumDataRows()+this.getFirstDataRow()-1;
		for (int i = 0; i <= lastRow; i++) {
			Row row = sheet.getRow(i);
			int lastCell = row.getLastCellNum();	// last cell + 1
			int maxNumLines = 1;
			for (int j = 0; j < lastCell; j++) {
				Cell cell = row.getCell(j);
				if (cell != null) {
					int cellLines = getNumWrappedLines(cell);
					if (cellLines > maxNumLines) {
						maxNumLines = cellLines;
					}
				}
			}
			if (maxNumLines > MAX_ROW_LINES) {
				maxNumLines = MAX_ROW_LINES;
			}
			if (maxNumLines > 1) {
				row.setHeight((short) (sheet.getDefaultRowHeight()*maxNumLines));
			}		
		}
	}

	/**
	 * @param cell
	 * @return
	 */
	private int getNumWrappedLines(Cell cell) {
		if (cell.getCellType() == CellType.STRING) {
			String val = cell.getStringCellValue();
			if (val == null || val.isEmpty()) {
				return 1;
			}
			CellStyle style = cell.getCellStyle();
			if (style == null || !style.getWrapText()) {
				return 1;
			}
			Font font = sheet.getWorkbook().getFontAt(style.getFontIndex());
			AttributedString astr = new AttributedString(val);
			java.awt.Font awtFont = new java.awt.Font(font.getFontName(), 0, font.getFontHeightInPoints());
			float cellWidth = sheet.getColumnWidth(cell.getColumnIndex())/ 256F * 5.5F;
			astr.addAttribute(TextAttribute.FONT, awtFont);
			FontRenderContext context = new FontRenderContext(null, true, true);
			java.awt.font.LineBreakMeasurer measurer = new java.awt.font.LineBreakMeasurer(astr.getIterator(), context);
			int pos = 0;
			int numLines = 0;
			while (measurer.getPosition() < val.length()) {
				pos = measurer.nextOffset(cellWidth);
				numLines++;
				measurer.setPosition(pos);
			}
			return numLines;
		} else {	// Not a string type
			return 1;
		}
	}
	
	/**
	 * Create a string from a collection of checksums
	 * @param checksumCollection collection of checksums
	 * @return string representation of the checksum
	 * @throws InvalidSPDXAnalysisException on SPDX parsing errors
	 */
	public String checksumsToString(Collection<Checksum> checksumCollection) throws InvalidSPDXAnalysisException {
		if (checksumCollection == null || checksumCollection.size() == 0) {
			return "";
		}
		Checksum[] checksums = checksumCollection.toArray(new Checksum[checksumCollection.size()]);
		Arrays.sort(checksums);
		StringBuilder sb = new StringBuilder(checksumToString(checksums[0]));
		for (int i = 1; i < checksums.length; i++) {
			sb.append("\n");
			String checksum = checksumToString(checksums[i]);
			sb.append(checksum);
		}
		return sb.toString();
	}
	

	/**
	 * @param checksum
	 * @return
	 * @throws InvalidSPDXAnalysisException 
	 */
	public String checksumToString(Checksum checksum) throws InvalidSPDXAnalysisException {
		if (checksum == null) {
			return "";
		}
		StringBuilder sb = new StringBuilder(checksum.getAlgorithm().toString().replaceAll("_", "-"));
		sb.append(": ");
		sb.append(checksum.getValue());
		return sb.toString();
	}
	
	/**
	 * @param checksumsString
	 * @return
	 * @throws InvalidSPDXAnalysisException 
	 */
	public Collection<Checksum> strToChecksums(String checksumsString) throws SpreadsheetException {
		if (checksumsString == null || checksumsString.trim().isEmpty()) {
			return new ArrayList<Checksum>();
		}
		String[] parts = checksumsString.split("\n");
		ArrayList<Checksum> retval = new ArrayList<>();
		for (int i = 0; i < parts.length; i++) {
			retval.add(parseChecksum(parts[i].trim()));
		}
		return retval;
	}
	
	/**
	 * Creates a Checksum from the parameters specified in the tag value
	 * @param value checksum string formatted with the algorithm
	 * @return Checksum
	 * @throws SpreadsheetException on errors parsing the checksum
	 */
	public Checksum parseChecksum(String value) throws SpreadsheetException {
		Matcher matcher = CHECKSUM_PATTERN.matcher(value.trim());
		if (!matcher.find()) {
			throw(new SpreadsheetException("Invalid checksum: "+value));
		}
		ChecksumAlgorithm algorithm;
		try {
			algorithm = ChecksumAlgorithm.valueOf(matcher.group(1).replaceAll("-", "_"));
		} catch (Exception ex) {
			algorithm = null;
		}
		if (algorithm == null) {
			throw(new SpreadsheetException("Invalid checksum algorithm: "+value));
		}
		try {
			return Checksum.create(modelStore, documentUri, algorithm, matcher.group(2));
		} catch (InvalidSPDXAnalysisException e) {
			throw(new SpreadsheetException("Error creating checksum for "+value,e));
		}
	}
	
	/**
	 * converts an array of strings to a comma separated list
	 * @param strings
	 * @return
	 */
	public static String stringsToCsv(Collection<String> strings) {
		StringWriter writer = new StringWriter();
		CSVWriter csvWriter = new CSVWriter(writer, CSV_SEPARATOR_CHAR, CSV_QUOTING_CHAR,
				CSV_ESCAPE_CHAR, CSV_LINE_END);
		try {
			csvWriter.writeNext(strings.toArray(new String[strings.size()]));
			csvWriter.flush();
			String retval = writer.toString().trim();
			return retval;
		} catch (Exception e) {
			return "ERROR PARSING CSV Entries";
		} finally {
			try {
				csvWriter.close();
			} catch (IOException e) {
				// ignore the close errors
			}
		}
	}
	
	/**
	 * Converts a comma separated CSV string to an array of strings
	 * @param csv
	 * @return
	 */
	public static List<String> csvToStrings(String csv) {
		StringReader reader = new StringReader(csv);
		final CSVReader csvReader = new CSVReaderBuilder(reader)
				.withCSVParser(parser)
				.build();
		try {
			return Arrays.asList(csvReader.readNext());
		} catch (IOException e) {
			return Arrays.asList(new String[] {"I/O ERROR PARSING CSV String"});
		} catch (CsvValidationException e) {
			return Arrays.asList(new String[] {"CSV VALIDATION ERROR PARSING CSV String"});
		} finally {
			try {
				csvReader.close();
			} catch (IOException e) {
				// Ignore
			}
		}
	}
	
	public static String licensesToString(Collection<AnyLicenseInfo> licenseCollection) {
		if (licenseCollection == null || licenseCollection.isEmpty()) {
			return "";
		}
		AnyLicenseInfo[] licenses = licenseCollection.toArray(new AnyLicenseInfo[licenseCollection.size()]);
		if (licenses.length == 1) {
			return licenses[0].toString();
		} else {
			StringBuilder sb = new StringBuilder(licenses[0].toString());
			for (int i = 1; i < licenses.length; i++) {
				sb.append(", ");
				sb.append(licenses[i].toString());
			}
			return sb.toString();
		}
	}
}
