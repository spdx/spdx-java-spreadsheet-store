/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.util.Iterator;
import java.util.NavigableMap;
import java.util.NoSuchElementException;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

/**
 * Adapter for Apache POI {@link Row} managing a specific row in a SODS {@link com.github.miachm.sods.Sheet}.
 */
public class OdsRow implements Row {
	private final OdsSheet sheet;
	private final int rowNum;
	private final NavigableMap<Integer, OdsCell> cells = new TreeMap<>();

	/**
	 * Creates an ODS row wrapper around a specific row index in a sheet.
	 *
	 * @param sheet Parent ODS sheet adapter.
	 * @param rowNum Zero-based row index.
	 */
	public OdsRow(OdsSheet sheet, int rowNum) {
		this.sheet = sheet;
		this.rowNum = rowNum;
	}

	/**
	 * Clears all cells in this row, setting their SODS range values to null.
	 */
	public void clear() {
		for (OdsCell cell : cells.values()) {
			cell.getSodsRange().setValue(null);
		}
		cells.clear();
	}

	private static final double MM_PER_PT = 25.4 / 72.0;

	@Override
	public Cell createCell(int column) {
		return createCell(column, CellType.BLANK);
	}

	@Override
	public synchronized Cell createCell(int column, CellType type) {
		if (column < 0) {
			throw new IllegalArgumentException("Column index must be >= 0");
		}
		com.github.miachm.sods.Sheet sodsSheet = sheet.getSodsSheet();
		int currentRows = sodsSheet.getMaxRows();
		if (rowNum >= currentRows) {
			sodsSheet.appendRows(rowNum - currentRows + 1);
		}
		int currentCols = sodsSheet.getMaxColumns();
		if (column >= currentCols) {
			sodsSheet.appendColumns(column - currentCols + 1);
		}
		com.github.miachm.sods.Range range = sodsSheet.getRange(rowNum, column);
		OdsCell cell = new OdsCell(this, column, range);
		if (type != null) {
			cell.setCellType(type);
		}
		cells.put(column, cell);
		return cell;
	}

	@Override
	public synchronized Cell getCell(int cellnum) {
		OdsCell cached = cells.get(cellnum);
		if (cached != null) {
			return cached;
		}
		com.github.miachm.sods.Sheet sodsSheet = sheet.getSodsSheet();
		if (cellnum >= 0 && cellnum < sodsSheet.getMaxColumns() && rowNum >= 0 && rowNum < sodsSheet.getMaxRows()) {
			com.github.miachm.sods.Range range = sodsSheet.getRange(rowNum, cellnum);
			if (range.getValue() != null || range.getFormula() != null || range.getAnnotation() != null
					|| (range.getStyle() != null && !range.getStyle().equals(new com.github.miachm.sods.Style()))) {
				OdsCell cell = new OdsCell(this, cellnum, range);
				cells.put(cellnum, cell);
				return cell;
			}
		}
		return null;
	}

	@Override
	public short getFirstCellNum() {
		com.github.miachm.sods.Range dataRange = sheet.getSodsSheet().getDataRange();
		if (!cells.isEmpty()) {
			int col = cells.firstKey();
			if (dataRange != null && rowNum >= dataRange.getRow() && rowNum <= dataRange.getLastRow()) {
				col = Math.min(col, dataRange.getColumn());
			}
			return col > Short.MAX_VALUE ? Short.MAX_VALUE : (short) col;
		}
		if (dataRange != null && rowNum >= dataRange.getRow() && rowNum <= dataRange.getLastRow()) {
			int col = dataRange.getColumn();
			return col > Short.MAX_VALUE ? Short.MAX_VALUE : (short) col;
		}
		return -1;
	}

	@Override
	public short getLastCellNum() {
		com.github.miachm.sods.Range dataRange = sheet.getSodsSheet().getDataRange();
		if (!cells.isEmpty()) {
			int nextCol = cells.lastKey() + 1;
			if (dataRange != null && rowNum >= dataRange.getRow() && rowNum <= dataRange.getLastRow()) {
				nextCol = Math.max(nextCol, dataRange.getLastColumn() + 1);
			}
			return nextCol > Short.MAX_VALUE ? Short.MAX_VALUE : (short) nextCol;
		}
		if (dataRange != null && rowNum >= dataRange.getRow() && rowNum <= dataRange.getLastRow()) {
			int nextCol = dataRange.getLastColumn() + 1;
			return nextCol > Short.MAX_VALUE ? Short.MAX_VALUE : (short) nextCol;
		}
		return -1;
	}

	@Override
	public int getRowNum() {
		return rowNum;
	}

	@Override
	public void setHeightInPoints(float height) {
		int currentRows = sheet.getSodsSheet().getMaxRows();
		if (rowNum >= currentRows) {
			sheet.getSodsSheet().appendRows(rowNum - currentRows + 1);
		}
		// Convert POI points to SODS millimeters (1 pt = 25.4 mm / 72 pt).
		sheet.getSodsSheet().setRowHeight(rowNum, height * MM_PER_PT);
	}

	@Override
	public void setHeight(short height) {
		setHeightInPoints(height / 20.0f);
	}

	@Override
	public float getHeightInPoints() {
		Double heightMm = sheet.getSodsSheet().getRowHeight(rowNum);
		if (heightMm != null) {
			// Convert SODS millimeters back to POI points.
			return (float) (heightMm / MM_PER_PT);
		}
		return 15.0f;
	}

	@Override
	public short getHeight() {
		return (short) Math.round(getHeightInPoints() * 20.0f);
	}
	@Override
	public Sheet getSheet() {
		return sheet;
	}
	@Override
	public Iterator<Cell> cellIterator() {
		return new Iterator<Cell>() {
			private int curCol = getFirstCellNum();
			private final int lastCol = getLastCellNum();
			private Cell nextCell = null;

			private void advance() {
				nextCell = null;
				if (curCol < 0) return;
				while (curCol >= 0 && curCol < lastCol) {
					Cell c = getCell(curCol);
					curCol++;
					if (c != null) {
						nextCell = c;
						break;
					}
				}
			}

			{
				advance();
			}

			@Override
			public boolean hasNext() {
				return nextCell != null;
			}

			@Override
			public Cell next() {
				if (!hasNext()) {
					throw new NoSuchElementException();
				}
				Cell res = nextCell;
				advance();
				return res;
			}

			@Override
			public void remove() {
				throw new UnsupportedOperationException("Remove not supported on cell iterator");
			}
		};
	}
	@Override
	public Iterator<Cell> iterator() {
		return cellIterator();
	}
	@Override
	public void removeCell(Cell cell) {
		if (cell instanceof OdsCell) {
			int col = cell.getColumnIndex();
			((OdsCell) cell).getSodsRange().setValue(null);
			cells.remove(col);
		}
	}
	@Override
	public void setRowNum(int rowNum) {}
	@Override
	public int getPhysicalNumberOfCells() { return cells.size(); }
	@Override
	public boolean isFormatted() { return false; }
	@Override
	public CellStyle getRowStyle() { return null; }
	@Override
	public void setRowStyle(CellStyle style) {}

	@Override
	public void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {}

	@Override
	public void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {}

	@Override
	public int getOutlineLevel() { return 0; }

	@Override
	public boolean getZeroHeight() {
		return sheet.getSodsSheet().rowIsHidden(rowNum);
	}

	@Override
	public void setZeroHeight(boolean zHeight) {
		int currentRows = sheet.getSodsSheet().getMaxRows();
		if (rowNum >= currentRows) {
			sheet.getSodsSheet().appendRows(rowNum - currentRows + 1);
		}
		if (zHeight) {
			sheet.getSodsSheet().hideRow(rowNum);
		} else {
			sheet.getSodsSheet().showRow(rowNum);
		}
	}

	@Override
	public Cell getCell(int cellnum, MissingCellPolicy policy) {
		if (policy == null) {
			policy = sheet.getWorkbook().getMissingCellPolicy();
		}
		Cell cell = getCell(cellnum);
		if (policy == MissingCellPolicy.RETURN_NULL_AND_BLANK) {
			return cell;
		}
		if (policy == MissingCellPolicy.RETURN_BLANK_AS_NULL) {
			if (cell == null || cell.getCellType() == CellType.BLANK) {
				return null;
			}
			return cell;
		}
		if (policy == MissingCellPolicy.CREATE_NULL_AS_BLANK) {
			if (cell == null) {
				return createCell(cellnum, CellType.BLANK);
			}
			return cell;
		}
		return cell;
	}
}
