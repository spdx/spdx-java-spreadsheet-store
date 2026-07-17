/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import com.github.miachm.sods.SpreadSheet;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;

/**
 * Adapter for Apache POI {@link Workbook} over a SODS {@link com.github.miachm.sods.SpreadSheet}.
 * Serves as the primary entry point for ODS read/write operations within the POI interface layer.
 */
public class OdsWorkbook implements Workbook {

	private final SpreadSheet spreadSheet;
	private final List<OdsSheet> sheets = new ArrayList<>();
	private final List<OdsFont> fonts = new ArrayList<>();
	private final List<OdsCellStyle> styles = new ArrayList<>();
	private final OdsCreationHelper creationHelper = new OdsCreationHelper(this);

	public OdsWorkbook() {
		this.spreadSheet = new SpreadSheet();
	}

	public OdsWorkbook(InputStream is) throws IOException {
		this.spreadSheet = new SpreadSheet(is);
		for (com.github.miachm.sods.Sheet sodsSheet : spreadSheet.getSheets()) {
			sheets.add(new OdsSheet(this, sodsSheet));
		}
	}

	public SpreadSheet getSpreadSheet() {
		return this.spreadSheet;
	}

	@Override
	public int getActiveSheetIndex() {
		return 0;
	}

	@Override
	public void setActiveSheet(int sheetIndex) {
	}

	@Override
	public int getFirstVisibleTab() {
		return 0;
	}

	@Override
	public void setFirstVisibleTab(int sheetIndex) {
	}

	@Override
	public void setSheetOrder(String sheetname, int pos) {
	}

	@Override
	public void setSelectedTab(int index) {
	}

	@Override
	public void setSheetName(int sheet, String name) {
		if (sheet >= 0 && sheet < sheets.size()) {
			sheets.get(sheet).getSodsSheet().setName(name);
		}
	}

	@Override
	public String getSheetName(int sheet) {
		if (sheet >= 0 && sheet < sheets.size()) {
			return sheets.get(sheet).getSheetName();
		}
		return null;
	}

	@Override
	public int getSheetIndex(String name) {
		for (int i = 0; i < sheets.size(); i++) {
			if (sheets.get(i).getSheetName().equalsIgnoreCase(name)) {
				return i;
			}
		}
		return -1;
	}

	@Override
	public int getSheetIndex(Sheet sheet) {
		return sheets.indexOf(sheet);
	}

	@Override
	public Sheet createSheet() {
		return createSheet("Sheet" + sheets.size());
	}

	@Override
	public Sheet createSheet(String sheetname) {
		com.github.miachm.sods.Sheet sodsSheet = new com.github.miachm.sods.Sheet(sheetname);
		spreadSheet.appendSheet(sodsSheet);
		OdsSheet newSheet = new OdsSheet(this, sodsSheet);
		sheets.add(newSheet);
		return newSheet;
	}

	@Override
	public Sheet cloneSheet(int sheetNum) {
		throw new UnsupportedOperationException("OdsWorkbook: cloneSheet not implemented");
	}

	@Override
	public Iterator<Sheet> sheetIterator() {
		List<Sheet> list = new ArrayList<>(sheets);
		return list.iterator();
	}

	@Override
	public int getNumberOfSheets() {
		return sheets.size();
	}

	@Override
	public Sheet getSheetAt(int index) {
		if (index >= 0 && index < sheets.size()) {
			return sheets.get(index);
		}
		return null;
	}

	@Override
	public Sheet getSheet(String name) {
		int index = getSheetIndex(name);
		return index != -1 ? getSheetAt(index) : null;
	}

	@Override
	public void removeSheetAt(int index) {
		if (index >= 0 && index < sheets.size()) {
			spreadSheet.deleteSheet(index);
			sheets.remove(index);
		}
	}

	@Override
	public Font createFont() {
		OdsFont font = new OdsFont((short) fonts.size());
		fonts.add(font);
		return font;
	}

	@Override
	public Font findFont(boolean bold, short color, short fontHeight, String name, boolean italic, boolean strikeout, short typeOffset, byte underline) {
		for (OdsFont font : fonts) {
			if (font.getBold() == bold && font.getColor() == color && font.getFontHeight() == fontHeight && font.getFontName().equals(name)) {
				return font;
			}
		}
		return null;
	}

	@Override
	public int getNumberOfFonts() {
		return fonts.size();
	}

	@Override
	public Font getFontAt(int idx) {
		if (idx >= 0 && idx < fonts.size()) {
			return fonts.get(idx);
		}
		return null;
	}

	@Override
	public DataFormat createDataFormat() {
		return creationHelper.createDataFormat();
	}

	@Override
	public CellStyle createCellStyle() {
		OdsCellStyle style = new OdsCellStyle(this);
		styles.add(style);
		return style;
	}

	@Override
	public int getNumCellStyles() {
		return styles.size();
	}

	@Override
	public CellStyle getCellStyleAt(int idx) {
		if (idx >= 0 && idx < styles.size()) {
			return styles.get(idx);
		}
		return null;
	}

	@Override
	public void write(OutputStream stream) throws IOException {
		spreadSheet.save(stream);
	}

	@Override
	public void close() throws IOException {
	}

	@Override
	public int addPicture(byte[] pictureData, int format) {
		return 0;
	}

	@Override
	public List<? extends PictureData> getAllPictures() {
		return new ArrayList<>();
	}

	@Override
	public CreationHelper getCreationHelper() {
		return creationHelper;
	}

	@Override
	public boolean isHidden() {
		return false;
	}

	@Override
	public void setHidden(boolean hidden) {
	}

	@Override
	public boolean isSheetHidden(int sheetNum) {
		return false;
	}

	@Override
	public boolean isSheetVeryHidden(int sheetNum) {
		return false;
	}

	@Override
	public void setSheetHidden(int sheetNum, boolean hidden) {
	}

	@Override
	public void setSheetVisibility(int sheetNum, SheetVisibility visibility) {
	}

	@Override
	public SheetVisibility getSheetVisibility(int sheetNum) {
		return SheetVisibility.VISIBLE;
	}

	@Override
	public void addToolPack(org.apache.poi.ss.formula.udf.UDFFinder toopack) {
	}

	@Override
	public void setForceFormulaRecalculation(boolean value) {
	}

	@Override
	public boolean getForceFormulaRecalculation() {
		return false;
	}

	@Override
	public SpreadsheetVersion getSpreadsheetVersion() {
		return SpreadsheetVersion.EXCEL2007;
	}

	@Override
	public Iterator<Sheet> iterator() {
		return sheetIterator();
	}

	@Override
	public void setCellReferenceType(org.apache.poi.ss.usermodel.CellReferenceType cellReferenceType) {
	}

	@Override
	public org.apache.poi.ss.usermodel.CellReferenceType getCellReferenceType() {
		return org.apache.poi.ss.usermodel.CellReferenceType.A1;
	}

	@Override
	public int addOlePackage(byte[] oleData, String label, String fileName, String command) throws IOException {
		return 0;
	}

	@Override
	public org.apache.poi.ss.formula.EvaluationWorkbook createEvaluationWorkbook() {
		return null;
	}

	@Override
	public void setMissingCellPolicy(Row.MissingCellPolicy missingCellPolicy) {}

	@Override
	public Row.MissingCellPolicy getMissingCellPolicy() {
		return Row.MissingCellPolicy.RETURN_NULL_AND_BLANK;
	}

	@Override
	public void removePrintArea(int sheetIndex) {}

	@Override
	public void setPrintArea(int sheetIndex, String reference) {}

	@Override
	public void setPrintArea(int sheetIndex, int startColumn, int endColumn, int startRow, int endRow) {}

	@Override
	public String getPrintArea(int sheetIndex) { return null; }

	@Override
	public int linkExternalWorkbook(String name, Workbook workbook) {
		return 0;
	}

	@Override
	public int getNumberOfNames() { return 0; }

	@Override
	public Name getName(String name) { return null; }

	@Override
	public List<? extends Name> getNames(String name) { return new ArrayList<>(); }

	@Override
	public List<? extends Name> getAllNames() { return new ArrayList<>(); }

	@Override
	public Name createName() { return null; }

	@Override
	public void removeName(Name name) {}

	@Override
	public int getNumberOfFontsAsInt() {
		return fonts.size();
	}
}
