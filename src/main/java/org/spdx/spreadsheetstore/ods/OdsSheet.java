/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.util.*;
import org.apache.poi.ss.usermodel.*;
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

	public OdsSheet(OdsWorkbook workbook, com.github.miachm.sods.Sheet sodsSheet) {
		this.workbook = workbook;
		this.sodsSheet = sodsSheet;
		for (int r = 0; r < sodsSheet.getMaxRows(); r++) {
			rows.put(r, new OdsRow(this, r));
		}
	}

	public com.github.miachm.sods.Sheet getSodsSheet() {
		return this.sodsSheet;
	}

	@Override
	public Row createRow(int rownum) {
		int currentRows = sodsSheet.getMaxRows();
		if (rownum >= currentRows) {
			for (int i = currentRows; i <= rownum; i++) {
				sodsSheet.appendRow();
			}
		}
		OdsRow row = new OdsRow(this, rownum);
		rows.put(rownum, row);
		return row;
	}

	@Override
	public Row getRow(int rownum) {
		return rows.get(rownum);
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

	@Override
	public int getFirstRowNum() {
		return rows.isEmpty() ? 0 : rows.keySet().iterator().next();
	}

	@Override
	public int getLastRowNum() {
		return rows.isEmpty() ? 0 : rows.lastKey();
	}

	@Override
	public void setColumnWidth(int columnIndex, int width) {
	}

	@Override
	public int getColumnWidth(int columnIndex) {
		return 2048;
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
	public void removeMergedRegion(int index) {}
	@Override
	public void removeMergedRegions(Collection<Integer> indices) {}
	@Override
	public int getNumMergedRegions() { return 0; }
	@Override
	public CellRangeAddress getMergedRegion(int index) { return null; }
	@Override
	public List<CellRangeAddress> getMergedRegions() { return new ArrayList<>(); }
	@Override
	public Iterator<Row> rowIterator() {
		List<Row> list = new ArrayList<>(rows.values());
		return list.iterator();
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
	public int addMergedRegion(CellRangeAddress region) { return 0; }
	@Override
	public int addMergedRegionUnsafe(CellRangeAddress region) { return 0; }

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
		return new HashMap<>();
	}

	@Override
	public Comment getCellComment(CellAddress address) { return null; }

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

	@Override
	public boolean isDisplayFormulas() { return false; }

	@Override
	public org.apache.poi.ss.util.PaneInformation getPaneInformation() { return null; }

	@Override
	public void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) {}

	@Override
	public void createFreezePane(int colSplit, int rowSplit) {}

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
	public void protectSheet(String password) {}

	@Override
	public boolean getProtect() { return false; }

	@Override
	public double getMargin(short margin) { return 0; }

	@Override
	public void setMargin(short margin, double size) {}

	@Override
	public void setSelected(boolean sel) {}

	@Override
	public Header getHeader() { return null; }

	@Override
	public Footer getFooter() { return null; }

	@Override
	public PrintSetup getPrintSetup() { return null; }

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
	public boolean isColumnHidden(int columnIndex) { return false; }

	@Override
	public void setColumnHidden(int columnIndex, boolean hidden) {}
}
