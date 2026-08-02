/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.NavigableMap;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.AutoFilter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.PageMargin;
import org.apache.poi.ss.usermodel.PaneType;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Adapter for Apache POI {@link Sheet} over a SODS {@link com.github.miachm.sods.Sheet}.
 * Manages rows, columns, and sheet-level properties for an ODS document.
 */
public class OdsSheet implements Sheet {

	private final OdsWorkbook workbook;
	private final com.github.miachm.sods.Sheet sodsSheet;
	private final NavigableMap<Integer, OdsRow> rows = new TreeMap<>();

	private final List<CellRangeAddress> mergedRegions = new ArrayList<>();

	/**
	 * Creates an ODS sheet wrapper around a SODS sheet.
	 *
	 * @param workbook Parent ODS workbook adapter.
	 * @param sodsSheet Underlying SODS sheet instance.
	 */
	public OdsSheet(OdsWorkbook workbook, com.github.miachm.sods.Sheet sodsSheet) {
		this.workbook = workbook;
		this.sodsSheet = sodsSheet;
	}

	/**
	 * Returns the underlying SODS sheet instance.
	 *
	 * @return The underlying SODS sheet.
	 */
	com.github.miachm.sods.Sheet getSodsSheet() {
		return this.sodsSheet;
	}

	@Override
	public synchronized Row createRow(int rownum) {
		if (rownum < 0) {
			throw new IllegalArgumentException("Row number must be >= 0");
		}
		int currentRows = sodsSheet.getMaxRows();
		if (rownum >= currentRows) {
			sodsSheet.appendRows(rownum - currentRows + 1);
		}
		OdsRow row = new OdsRow(this, rownum);
		rows.put(rownum, row);
		return row;
	}

	@Override
	public synchronized Row getRow(int rownum) {
		OdsRow row = rows.get(rownum);
		if (row != null) {
			return row;
		}
		if (rownum >= 0 && rownum < sodsSheet.getMaxRows()) {
			OdsRow newRow = new OdsRow(this, rownum);
			rows.put(rownum, newRow);
			return newRow;
		}
		return null;
	}

	@Override
	public void removeRow(Row row) {
		if (row instanceof OdsRow) {
			int rowNum = row.getRowNum();
			OdsRow odsRow = (OdsRow) row;
			odsRow.clear();
			rows.remove(rowNum);
		}
	}

	private boolean hasData() {
		if (!rows.isEmpty()) {
			return true;
		}
		com.github.miachm.sods.Range dataRange = sodsSheet.getDataRange();
		if (dataRange == null) {
			return false;
		}
		for (int r = dataRange.getRow(); r <= dataRange.getLastRow(); r++) {
			for (int c = dataRange.getColumn(); c <= dataRange.getLastColumn(); c++) {
				com.github.miachm.sods.Range range = sodsSheet.getRange(r, c);
				if (range.getValue() != null || range.getFormula() != null || range.getAnnotation() != null) {
					return true;
				}
			}
		}
		return false;
	}

	@Override
	public int getFirstRowNum() {
		if (!hasData()) {
			return -1;
		}
		com.github.miachm.sods.Range dataRange = sodsSheet.getDataRange();
		if (!rows.isEmpty()) {
			int firstRow = rows.firstKey();
			if (dataRange != null) {
				return Math.min(firstRow, dataRange.getRow());
			}
			return firstRow;
		}
		return dataRange.getRow();
	}

	@Override
	public int getLastRowNum() {
		if (!hasData()) {
			return -1;
		}
		com.github.miachm.sods.Range dataRange = sodsSheet.getDataRange();
		if (!rows.isEmpty()) {
			int lastRow = rows.lastKey();
			if (dataRange != null) {
				return Math.max(lastRow, dataRange.getLastRow());
			}
			return lastRow;
		}
		return dataRange.getLastRow();
	}

	@Override
	public void setColumnWidth(int columnIndex, int width) {
		if (columnIndex < 0) {
			throw new IllegalArgumentException("Column index must be >= 0");
		}
		int currentCols = sodsSheet.getMaxColumns();
		if (columnIndex >= currentCols) {
			sodsSheet.appendColumns(columnIndex - currentCols + 1);
		}
		// Convert POI 1/256th character units to SODS millimeters.
		// 1 character width (e.g. 10pt Arial '0') = 5.25 pt = 5.25 * (25.4/72) mm ≈ 1.852 mm.
		double widthMm = (width / 256.0) * 1.852;
		sodsSheet.setColumnWidth(columnIndex, widthMm);
	}

	@Override
	public int getColumnWidth(int columnIndex) {
		if (columnIndex >= 0 && columnIndex < sodsSheet.getMaxColumns()) {
			double widthMm = sodsSheet.getColumnWidth(columnIndex);
			if (widthMm > 0) {
				// Convert SODS millimeters back to POI 1/256th character units.
				return (int) Math.round((widthMm / 1.852) * 256.0);
			}
		}
		return 2048; // default width
	}

	@Override
	public void autoSizeColumn(int columnIndex) {
	}

	@Override
	public Workbook getWorkbook() {
		return workbook;
	}

	@Override
	public String getSheetName() {
		return sodsSheet.getName();
	}

	@Override
	public double getMargin(PageMargin margin) { return 0; }
	@Override
	public void setMargin(PageMargin margin, double size) {}
	@Override
	public boolean isPrintGridlines() { return false; }
	@Override
	public void setPrintGridlines(boolean show) {}
	@Override
	public boolean isPrintRowAndColumnHeadings() { return false; }
	@Override
	public void setPrintRowAndColumnHeadings(boolean show) {}
	@Override
	public void setActiveCell(CellAddress address) {}
	@Override
	public void removeMergedRegion(int index) {
		if (index < 0 || index >= mergedRegions.size()) {
			throw new IllegalArgumentException("Invalid merged region index: " + index);
		}
		CellRangeAddress region = mergedRegions.remove(index);
		int firstRow = region.getFirstRow();
		int firstCol = region.getFirstColumn();
		int numRows = region.getLastRow() - firstRow + 1;
		int numCols = region.getLastColumn() - firstCol + 1;
		com.github.miachm.sods.Range range = sodsSheet.getRange(firstRow, firstCol, numRows, numCols);
		range.split();
	}

	@Override
	public void removeMergedRegions(Collection<Integer> indices) {
		if (indices != null) {
			List<Integer> sorted = new ArrayList<>();
			for (Integer idx : indices) {
				if (idx != null) {
					sorted.add(idx);
				}
			}
			Collections.sort(sorted, Collections.reverseOrder());
			for (int idx : sorted) {
				if (idx >= 0 && idx < mergedRegions.size()) {
					removeMergedRegion(idx);
				}
			}
		}
	}

	@Override
	public int getNumMergedRegions() {
		return mergedRegions.size();
	}

	@Override
	public CellRangeAddress getMergedRegion(int index) {
		if (index < 0 || index >= mergedRegions.size()) {
			throw new IllegalArgumentException("Invalid merged region index: " + index);
		}
		return mergedRegions.get(index);
	}

	@Override
	public List<CellRangeAddress> getMergedRegions() {
		return new ArrayList<>(mergedRegions);
	}
	@Override
	public Iterator<Row> rowIterator() {
		return new Iterator<Row>() {
			private int curRow = getFirstRowNum();
			private final int lastRow = getLastRowNum();
			private Row nextRow = null;

			private void advance() {
				nextRow = null;
				if (curRow < 0) return;
				while (curRow <= lastRow) {
					Row r = getRow(curRow);
					curRow++;
					if (r != null) {
						nextRow = r;
						break;
					}
				}
			}

			{
				advance();
			}

			@Override
			public boolean hasNext() {
				return nextRow != null;
			}

			@Override
			public Row next() {
				if (!hasNext()) {
					throw new java.util.NoSuchElementException();
				}
				Row res = nextRow;
				advance();
				return res;
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException("Remove not supported on row iterator");
			}
		};
	}
	@Override
	public Iterator<Row> iterator() {
		return rowIterator();
	}
	@Override
	public void setForceFormulaRecalculation(boolean value) {}
	@Override
	public boolean getForceFormulaRecalculation() { return false; }
	@Override
	public void setAutobreaks(boolean value) {}
	@Override
	public boolean getAutobreaks() { return false; }
	@Override
	public void setDisplayGuts(boolean value) {}
	@Override
	public boolean getDisplayGuts() { return false; }
	@Override
	public void setDisplayRowColHeadings(boolean value) {}
	@Override
	public void setDisplayFormulas(boolean value) {}
	@Override
	public void setDisplayGridlines(boolean value) {}
	@Override
	public boolean isDisplayGridlines() { return false; }
	@Override
	public void setRowSumsBelow(boolean value) {}
	@Override
	public boolean getRowSumsBelow() { return false; }
	@Override
	public void setRowSumsRight(boolean value) {}
	@Override
	public boolean getRowSumsRight() { return false; }
	@Override
	public int getPhysicalNumberOfRows() { return rows.size(); }
	@Override
	public int addMergedRegion(CellRangeAddress region) {
		if (region == null) {
			throw new IllegalArgumentException("Merged region cannot be null");
		}
		int firstRow = region.getFirstRow();
		int lastRow = region.getLastRow();
		int firstCol = region.getFirstColumn();
		int lastCol = region.getLastColumn();
		if (lastRow >= sodsSheet.getMaxRows()) {
			sodsSheet.appendRows(lastRow - sodsSheet.getMaxRows() + 1);
		}
		if (lastCol >= sodsSheet.getMaxColumns()) {
			sodsSheet.appendColumns(lastCol - sodsSheet.getMaxColumns() + 1);
		}
		int numRows = lastRow - firstRow + 1;
		int numCols = lastCol - firstCol + 1;
		com.github.miachm.sods.Range range = sodsSheet.getRange(firstRow, firstCol, numRows, numCols);
		range.merge();
		mergedRegions.add(region);
		return mergedRegions.size() - 1;
	}

	@Override
	public int addMergedRegionUnsafe(CellRangeAddress region) {
		return addMergedRegion(region);
	}

	@Override
	public CellAddress getActiveCell() { return null; }

	@Override
	public List<? extends Hyperlink> getHyperlinkList() { return new ArrayList<>(); }

	@Override
	public Hyperlink getHyperlink(CellAddress address) { return null; }

	@Override
	public Hyperlink getHyperlink(int row, int column) { return null; }

	@Override
	public int getColumnOutlineLevel(int columnIndex) { return 0; }

	@Override
	public void setRepeatingColumns(CellRangeAddress repeatingColumnsToIndex) {}

	@Override
	public void setRepeatingRows(CellRangeAddress repeatingRowsToIndex) {}

	@Override
	public CellRangeAddress getRepeatingColumns() { return null; }

	@Override
	public CellRangeAddress getRepeatingRows() { return null; }

	@Override
	public SheetConditionalFormatting getSheetConditionalFormatting() { return null; }

	@Override
	public AutoFilter setAutoFilter(CellRangeAddress range) { return null; }

	@Override
	public void addValidationData(DataValidation dataValidation) {}

	@Override
	public List<? extends DataValidation> getDataValidations() { return new ArrayList<>(); }

	@Override
	public DataValidationHelper getDataValidationHelper() { return null; }

	@Override
	public CellRange<? extends Cell> removeArrayFormula(Cell cell) { return null; }

	@Override
	public CellRange<? extends Cell> setArrayFormula(String formula, CellRangeAddress range) { return null; }

	@Override
	public boolean isSelected() { return false; }

	@Override
	public Drawing<?> createDrawingPatriarch() { return null; }

	@Override
	public Drawing<?> getDrawingPatriarch() { return null; }

	@Override
	public Map<CellAddress, ? extends Comment> getCellComments() {
		Map<CellAddress, Comment> map = new HashMap<>();
		com.github.miachm.sods.Range dataRange = sodsSheet.getDataRange();
		if (dataRange == null) {
			return map;
		}
		int startRow = dataRange.getRow();
		int endRow = dataRange.getLastRow();
		int startCol = dataRange.getColumn();
		int endCol = dataRange.getLastColumn();
		for (int r = startRow; r <= endRow; r++) {
			for (int c = startCol; c <= endCol; c++) {
				com.github.miachm.sods.Range range = sodsSheet.getRange(r, c);
				if (range.getAnnotation() != null) {
					CellAddress addr = new CellAddress(r, c);
					OdsComment comment = new OdsComment(range.getAnnotation());
					comment.setAddress(addr);
					map.put(addr, comment);
				}
			}
		}
		return map;
	}

	@Override
	public Comment getCellComment(CellAddress address) {
		if (address == null) return null;
		int r = address.getRow();
		int c = address.getColumn();
		if (r >= 0 && r < sodsSheet.getMaxRows() && c >= 0 && c < sodsSheet.getMaxColumns()) {
			com.github.miachm.sods.Range range = sodsSheet.getRange(r, c);
			if (range.getAnnotation() != null) {
				OdsComment comment = new OdsComment(range.getAnnotation());
				comment.setAddress(address);
				return comment;
			}
		}
		return null;
	}

	@Override
	public void autoSizeColumn(int columnIndex, boolean useMergedCells) {}

	@Override
	public void setDefaultColumnStyle(int column, CellStyle style) {}

	@Override
	public void setRowGroupCollapsed(int row, boolean collapse) {}

	@Override
	public void groupRow(int startRow, int endRow) {}

	@Override
	public void ungroupRow(int startRow, int endRow) {}

	@Override
	public void groupColumn(int startColumn, int endColumn) {}

	@Override
	public void ungroupColumn(int startColumn, int endColumn) {}

	@Override
	public void setColumnGroupCollapsed(int columnNumber, boolean collapsed) {}

	@Override
	public void setRowBreak(int row) {}

	@Override
	public void removeRowBreak(int row) {}

	@Override
	public boolean isRowBroken(int row) { return false; }

	@Override
	public int[] getRowBreaks() { return new int[0]; }

	@Override
	public void setColumnBreak(int column) {}

	@Override
	public void removeColumnBreak(int column) {}

	@Override
	public boolean isColumnBroken(int column) { return false; }

	@Override
	public int[] getColumnBreaks() { return new int[0]; }

	@Override
	public boolean isDisplayRowColHeadings() { return true; }

	private boolean protectedState = false;
	private final Map<Short, Double> marginMap = new HashMap<>();
	private int freezeColSplit = 0;
	private int freezeRowSplit = 0;

	@Override
	public boolean isDisplayFormulas() { return false; }

	@Override
	public org.apache.poi.ss.util.PaneInformation getPaneInformation() {
		if (freezeColSplit > 0 || freezeRowSplit > 0) {
			return new org.apache.poi.ss.util.PaneInformation((short) freezeColSplit, (short) freezeRowSplit, (short) freezeRowSplit, (short) freezeColSplit, (byte) 0, true);
		}
		return null;
	}

	@Override
	public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) {
		createFreezePane(colSplit, rowSplit);
	}

	@Override
	public void createFreezePane(int colSplit, int rowSplit) {
		this.freezeColSplit = colSplit;
		this.freezeRowSplit = rowSplit;
		if (rowSplit > 0) {
			sodsSheet.freezeRows(rowSplit);
		}
		if (colSplit > 0) {
			sodsSheet.freezeColumns(colSplit);
		}
	}

	@Override
	public void createSplitPane(int xSplit, int ySplit, int leftmostColumn, int topRow, PaneType activePane) {}

	@Override
	public void createSplitPane(int xSplit, int ySplit, int leftmostColumn, int topRow, int activePane) {}

	@Override
	public void shiftColumns(int startColumn, int endColumn, int n) {}

	@Override
	public void shiftRows(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight) {}

	@Override
	public void shiftRows(int startRow, int endRow, int n) {}

	@Override
	public void showInPane(int topRow, int leftmostColumn) {}

	@Override
	public short getLeftCol() { return 0; }

	@Override
	public short getTopRow() { return 0; }

	@Override
	public void setZoom(int scale) {}

	@Override
	public boolean getScenarioProtect() { return false; }

	@Override
	public void protectSheet(String password) {
		this.protectedState = true;
		if (password != null && !password.isEmpty()) {
			try {
				sodsSheet.setPassword(password);
			} catch (Exception e) {
				throw new RuntimeException("Failed to set sheet password", e);
			}
		}
	}

	@Override
	public boolean getProtect() {
		return protectedState || sodsSheet.isProtected();
	}

	@Override
	public double getMargin(short margin) {
		Double val = marginMap.get(margin);
		return val != null ? val : 0.75;
	}

	@Override
	public void setMargin(short margin, double size) {
		marginMap.put(margin, size);
	}

	@Override
	public void setSelected(boolean sel) {}

	private static class OdsHeaderFooterStub implements Header, Footer {
		private String left = "";
		private String center = "";
		private String right = "";

		@Override public String getLeft() { return left; }
		@Override public void setLeft(String newLeft) { this.left = newLeft != null ? newLeft : ""; }
		@Override public String getCenter() { return center; }
		@Override public void setCenter(String newCenter) { this.center = newCenter != null ? newCenter : ""; }
		@Override public String getRight() { return right; }
		@Override public void setRight(String newRight) { this.right = newRight != null ? newRight : ""; }
	}

	private final OdsHeaderFooterStub headerStub = new OdsHeaderFooterStub();
	private final OdsHeaderFooterStub footerStub = new OdsHeaderFooterStub();

	@Override
	public Header getHeader() { return headerStub; }

	@Override
	public Footer getFooter() { return footerStub; }

	private static class OdsPrintSetupStub implements PrintSetup {
		private boolean landscape = false;
		private short paperSize = LETTER_PAPERSIZE;
		@Override public void setPaperSize(short size) { this.paperSize = size; }
		@Override public void setScale(short scale) {}
		@Override public void setPageStart(short start) {}
		@Override public void setFitWidth(short width) {}
		@Override public void setFitHeight(short height) {}
		@Override public void setLeftToRight(boolean ltr) {}
		@Override public void setLandscape(boolean ls) { this.landscape = ls; }
		@Override public void setValidSettings(boolean valid) {}
		@Override public void setNoColor(boolean mono) {}
		@Override public void setDraft(boolean draft) {}
		@Override public void setNotes(boolean printNotes) {}
		@Override public void setNoOrientation(boolean orientation) {}
		@Override public void setUsePage(boolean page) {}
		@Override public void setHResolution(short resolution) {}
		@Override public void setVResolution(short resolution) {}
		@Override public void setCopies(short copies) {}
		@Override public short getPaperSize() { return paperSize; }
		@Override public short getScale() { return 100; }
		@Override public short getPageStart() { return 1; }
		@Override public short getFitWidth() { return 1; }
		@Override public short getFitHeight() { return 1; }
		@Override public boolean getLeftToRight() { return false; }
		@Override public boolean getLandscape() { return landscape; }
		@Override public boolean getValidSettings() { return true; }
		@Override public boolean getNoColor() { return false; }
		@Override public boolean getDraft() { return false; }
		@Override public boolean getNotes() { return false; }
		@Override public boolean getNoOrientation() { return false; }
		@Override public boolean getUsePage() { return false; }
		@Override public short getHResolution() { return 300; }
		@Override public short getVResolution() { return 300; }
		@Override public void setHeaderMargin(double headmargin) {}
		@Override public void setFooterMargin(double footmargin) {}
		@Override public double getHeaderMargin() { return 0.5; }
		@Override public double getFooterMargin() { return 0.5; }
		@Override public short getCopies() { return 1; }
	}

	private final OdsPrintSetupStub printSetupStub = new OdsPrintSetupStub();

	@Override
	public PrintSetup getPrintSetup() { return printSetupStub; }

	@Override
	public boolean getFitToPage() { return false; }

	@Override
	public void setFitToPage(boolean value) {}

	@Override
	public boolean isDisplayZeros() { return true; }

	@Override
	public void setDisplayZeros(boolean value) {}

	@Override
	public boolean getHorizontallyCenter() { return false; }

	@Override
	public void setHorizontallyCenter(boolean value) {}

	@Override
	public boolean getVerticallyCenter() { return false; }

	@Override
	public void setVerticallyCenter(boolean value) {}

	@Override
	public void validateMergedRegions() {}

	@Override
	public CellStyle getColumnStyle(int column) { return null; }

	@Override
	public short getDefaultRowHeight() { return 300; }

	@Override
	public void setDefaultRowHeight(short height) {}

	@Override
	public float getDefaultRowHeightInPoints() { return 15.0f; }

	@Override
	public void setDefaultRowHeightInPoints(float height) {}

	@Override
	public int getDefaultColumnWidth() { return 8; }

	@Override
	public void setDefaultColumnWidth(int width) {}

	@Override
	public float getColumnWidthInPixels(int columnIndex) { return 8.0f * 8; }

	@Override
	public boolean isRightToLeft() { return false; }

	@Override
	public void setRightToLeft(boolean value) {}

	@Override
	public boolean isColumnHidden(int columnIndex) {
		if (columnIndex >= 0 && columnIndex < sodsSheet.getMaxColumns()) {
			return sodsSheet.columnIsHidden(columnIndex);
		}
		return false;
	}

	@Override
	public void setColumnHidden(int columnIndex, boolean hidden) {
		if (columnIndex >= 0) {
			int currentCols = sodsSheet.getMaxColumns();
			if (columnIndex >= currentCols) {
				sodsSheet.appendColumns(columnIndex - currentCols + 1);
			}
			if (hidden) {
				sodsSheet.hideColumn(columnIndex);
			} else {
				sodsSheet.showColumn(columnIndex);
			}
		}
	}
}
