/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.util.Date;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * Adapter for Apache POI {@link Cell} over the SODS {@link com.github.miachm.sods.Range}.
 * Provides mechanisms to read and write cell data, respecting SODS limitations.
 */
public class OdsCell implements Cell {
	private final OdsRow row;
	private final int columnIndex;
	private final com.github.miachm.sods.Range range;
	private CellStyle cellStyle;

	public OdsCell(OdsRow row, int columnIndex, com.github.miachm.sods.Range range) {
		this.row = row;
		this.columnIndex = columnIndex;
		this.range = range;
	}

	public com.github.miachm.sods.Range getSodsRange() {
		return range;
	}

	@Override
	public void setCellValue(String value) {
		range.setValue(value);
	}

	@Override
	public void setCellValue(double value) {
		range.setValue(value);
	}

	@Override
	public void setCellValue(Date value) {
		if (value == null) {
			range.setValue(null);
		} else {
			java.time.LocalDateTime ldt = java.time.LocalDateTime.ofInstant(value.toInstant(), java.time.ZoneId.systemDefault());
			range.setValue(ldt);
		}
	}

	@Override
	public void setCellType(CellType cellType) {
	}

	@Override
	public String getStringCellValue() {
		Object val = range.getValue();
		return val == null ? "" : val.toString();
	}

	@Override
	public double getNumericCellValue() {
		Object val = range.getValue();
		if (val instanceof Number) {
			return ((Number) val).doubleValue();
		}
		return 0.0;
	}

	@Override
	public Date getDateCellValue() {
		Object value = range.getValue();
		if (value == null) return null;
		if (value instanceof java.time.LocalDateTime) {
			java.time.LocalDateTime ldt = (java.time.LocalDateTime) value;
			return java.util.Date.from(ldt.atZone(java.time.ZoneId.systemDefault()).toInstant());
		}
		if (value instanceof java.time.LocalDate) {
			java.time.LocalDate ld = (java.time.LocalDate) value;
			return java.util.Date.from(ld.atStartOfDay(java.time.ZoneId.systemDefault()).toInstant());
		}
		if (value instanceof java.util.Date) {
			return (Date) value;
		}
		if (value instanceof Number) {
			return org.apache.poi.ss.usermodel.DateUtil.getJavaDate(((Number) value).doubleValue());
		}
		return null;
	}

	@Override
	public CellType getCellType() {
		Object val = range.getValue();
		if (val == null) return CellType.BLANK;
		if (val instanceof String) return CellType.STRING;
		if (val instanceof Number) return CellType.NUMERIC;
		if (val instanceof Boolean) return CellType.BOOLEAN;
		if (val instanceof java.time.LocalDateTime || val instanceof java.time.LocalDate || val instanceof java.util.Date) {
			return CellType.NUMERIC;
		}
		return CellType.STRING;
	}

	@Override
	public void setCellStyle(CellStyle style) {
		this.cellStyle = style;
		if (style instanceof OdsCellStyle) {
			range.setStyle(((OdsCellStyle) style).getSodsStyle());
		}
	}

	@Override
	public CellStyle getCellStyle() {
		return cellStyle;
	}

	@Override
	public int getColumnIndex() {
		return columnIndex;
	}

	@Override
	public int getRowIndex() {
		return row.getRowNum();
	}

	@Override
	public Sheet getSheet() {
		return row.getSheet();
	}

	@Override
	public Row getRow() {
		return row;
	}

	@Override
	public void setBlank() {
		range.setValue(null);
	}

	@Override
	public void setCellFormula(String formula) {
	}

	@Override
	public String getCellFormula() {
		return "";
	}

	@Override
	public boolean getBooleanCellValue() {
		Object val = range.getValue();
		if (val instanceof Boolean) {
			return (Boolean) val;
		}
		return false;
	}

	@Override
	public byte getErrorCellValue() {
		return 0;
	}

	@Override
	public void setCellErrorValue(byte value) {}
	@Override
	public void setAsActiveCell() {}
	@Override
	public CellAddress getAddress() {
		return new CellAddress(getRowIndex(), getColumnIndex());
	}
	@Override
	public void setCellValue(java.util.Calendar value) {
		if (value != null) {
			setCellValue(value.getTime());
		}
	}
	@Override
	public void setCellValue(RichTextString value) {
		if (value != null) {
			setCellValue(value.getString());
		}
	}
	@Override
	public void setCellValue(boolean value) {
		range.setValue(value);
	}
	@Override
	public RichTextString getRichStringCellValue() {
		return new org.apache.poi.xssf.usermodel.XSSFRichTextString(getStringCellValue());
	}
	@Override
	public void removeCellComment() {}
	@Override
	public Comment getCellComment() { return null; }
	@Override
	public void setCellComment(Comment comment) {}
	@Override
	public Hyperlink getHyperlink() { return null; }
	@Override
	public void setHyperlink(Hyperlink hyperlink) {}
	@Override
	public CellRangeAddress getArrayFormulaRange() { return null; }
	@Override
	public boolean isPartOfArrayFormulaGroup() { return false; }

	@Override
	public void removeHyperlink() {}

	@Override
	public java.time.LocalDateTime getLocalDateTimeCellValue() {
		Object value = range.getValue();
		if (value instanceof java.time.LocalDateTime) {
			return (java.time.LocalDateTime) value;
		}
		if (value instanceof java.time.LocalDate) {
			return ((java.time.LocalDate) value).atStartOfDay();
		}
		if (value instanceof Date) {
			return java.time.LocalDateTime.ofInstant(((Date) value).toInstant(), java.time.ZoneId.systemDefault());
		}
		if (value instanceof Number) {
			Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(((Number) value).doubleValue());
			return java.time.LocalDateTime.ofInstant(date.toInstant(), java.time.ZoneId.systemDefault());
		}
		return null;
	}

	@Override
	public void setCellValue(java.time.LocalDateTime value) {
		range.setValue(value);
	}

	@Override
	public void setCellValue(java.time.LocalDate value) {
		range.setValue(value);
	}

	@Override
	public void removeFormula() {}

	@Override
	public CellType getCachedFormulaResultType() { return CellType.BLANK; }
}
