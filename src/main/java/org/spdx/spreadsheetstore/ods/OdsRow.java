/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.util.*;
import org.apache.poi.ss.usermodel.*;

/**
 * Adapter for Apache POI {@link Row} managing a specific row in a SODS {@link com.github.miachm.sods.Sheet}.
 */
public class OdsRow implements Row {
	private final OdsSheet sheet;
	private final int rowNum;
	private final Map<Integer, OdsCell> cells = new TreeMap<>();

	public OdsRow(OdsSheet sheet, int rowNum) {
		this.sheet = sheet;
		this.rowNum = rowNum;
	}

	public void clear() {
		for (OdsCell cell : cells.values()) {
			cell.getSodsRange().setValue(null);
		}
		cells.clear();
	}

	@Override
	public Cell createCell(int column) {
		return createCell(column, CellType.BLANK);
	}

	@Override
	public Cell createCell(int column, CellType type) {
		com.github.miachm.sods.Sheet sodsSheet = sheet.getSodsSheet();
		int currentCols = sodsSheet.getMaxColumns();
		if (column >= currentCols) {
			for (int i = currentCols; i <= column; i++) {
				sodsSheet.appendColumn();
			}
		}
		com.github.miachm.sods.Range range = sodsSheet.getRange(rowNum, column);
		OdsCell cell = new OdsCell(this, column, range);
		cells.put(column, cell);
		return cell;
	}

	@Override
	public Cell getCell(int cellnum) {
		// If it exists in cells map, return it
		if (cells.containsKey(cellnum)) {
			return cells.get(cellnum);
		}
		// If it exists in the underlying SODS sheet but not yet wrapped, wrap and return
		com.github.miachm.sods.Sheet sodsSheet = sheet.getSodsSheet();
		if (cellnum >= 0 && cellnum < sodsSheet.getMaxColumns()) {
			com.github.miachm.sods.Range range = sodsSheet.getRange(rowNum, cellnum);
			Object val = range.getValue();
			if (val != null) {
				if (val instanceof String && ((String) val).trim().isEmpty()) {
					return null;
				}
				OdsCell cell = new OdsCell(this, cellnum, range);
				cells.put(cellnum, cell);
				return cell;
			}
		}
		return null;
	}

	@Override
	public short getFirstCellNum() {
		com.github.miachm.sods.Sheet sodsSheet = sheet.getSodsSheet();
		int maxCols = sodsSheet.getMaxColumns();
		for (int col = 0; col < maxCols; col++) {
			if (cells.containsKey(col)) {
				return (short) col;
			}
			Object val = sodsSheet.getRange(rowNum, col).getValue();
			if (val != null) {
				return (short) col;
			}
		}
		return -1;
	}

	@Override
	public short getLastCellNum() {
		com.github.miachm.sods.Sheet sodsSheet = sheet.getSodsSheet();
		int maxCols = sodsSheet.getMaxColumns();
		for (int col = maxCols - 1; col >= 0; col--) {
			if (cells.containsKey(col)) {
				return (short) (col + 1);
			}
			Object val = sodsSheet.getRange(rowNum, col).getValue();
			if (val != null) {
				return (short) (col + 1);
			}
		}
		return -1;
	}

	@Override
	public int getRowNum() {
		return rowNum;
	}

	@Override
	public void setHeight(short height) {}
	@Override
	public void setHeightInPoints(float height) {}
	@Override
	public short getHeight() { return 255; }
	@Override
	public float getHeightInPoints() { return 15.0f; }
	@Override
	public Sheet getSheet() {
		return sheet;
	}
	@Override
	public Iterator<Cell> cellIterator() {
		List<Cell> list = new ArrayList<>(cells.values());
		return list.iterator();
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
	public boolean getZeroHeight() { return false; }

	@Override
	public void setZeroHeight(boolean zHeight) {}

	@Override
	public Cell getCell(int cellnum, MissingCellPolicy policy) {
		Cell cell = getCell(cellnum);
		if (cell == null && policy == MissingCellPolicy.CREATE_NULL_AS_BLANK) {
			return createCell(cellnum);
		}
		return cell;
	}
}
