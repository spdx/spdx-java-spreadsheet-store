/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.util.Date;
import java.time.LocalDate;
import java.time.LocalDateTime;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Hyperlink;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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

	/**
	 * Creates an ODS cell wrapper.
	 *
	 * @param row Parent ODS row adapter.
	 * @param columnIndex Zero-based column index.
	 * @param range Underlying SODS Range representing this cell.
	 */
	public OdsCell(OdsRow row, int columnIndex, com.github.miachm.sods.Range range) {
		this.row = row;
		this.columnIndex = columnIndex;
		this.range = range;
	}

	/**
	 * Returns the underlying SODS Range instance.
	 *
	 * @return The underlying SODS Range.
	 */
	com.github.miachm.sods.Range getSodsRange() {
		return range;
	}

	@Override
	public void setCellValue(String value) {
		range.setFormula(null);
		range.setValue(value);
	}

	@Override
	public void setCellValue(double value) {
		range.setFormula(null);
		range.setValue(value);
	}

	@Override
	public void setCellValue(Date value) {
		range.setFormula(null);
		if (value == null) {
			range.setValue(null);
		} else {
			java.time.LocalDateTime ldt = java.time.LocalDateTime.ofInstant(value.toInstant(), java.time.ZoneOffset.UTC);
			range.setValue(ldt);
		}
	}

	@Override
	public void setCellType(CellType cellType) {
		if (cellType == CellType.BLANK) {
			range.setValue(null);
			range.setFormula(null);
		} else if (cellType == CellType.STRING) {
			Object val = range.getValue();
			if (val != null) {
				range.setValue(val.toString());
			}
		}
	}

	@Override
	public String getStringCellValue() {
		Object val = range.getValue();
		return val == null ? "" : val.toString();
	}

	@Override
	public double getNumericCellValue() {
		CellType type = getCellType();
		if (type != CellType.NUMERIC && type != CellType.FORMULA) {
			throw new IllegalStateException("Cannot get a NUMERIC value from a " + type + " cell");
		}
		Object val = range.getValue();
		if (val instanceof Number) {
			return ((Number) val).doubleValue();
		}
		if (val instanceof com.github.miachm.sods.OfficeCurrency) {
			Double d = ((com.github.miachm.sods.OfficeCurrency) val).getValue();
			return d != null ? d : 0.0;
		}
		if (val instanceof com.github.miachm.sods.OfficePercentage) {
			Double d = ((com.github.miachm.sods.OfficePercentage) val).getValue();
			return d != null ? d : 0.0;
		}
		if (val instanceof java.util.Date) {
			return DateUtil.getExcelDate((java.util.Date) val);
		}
		if (val instanceof java.time.LocalDateTime) {
			return DateUtil.getExcelDate(java.sql.Timestamp.valueOf((java.time.LocalDateTime) val));
		}
		if (val instanceof java.time.LocalDate) {
			return DateUtil.getExcelDate(java.sql.Date.valueOf((java.time.LocalDate) val));
		}
		return 0.0;
	}

	@Override
	public Date getDateCellValue() {
		Object value = range.getValue();
		if (value == null) return null;
		if (value instanceof LocalDateTime) {
			LocalDateTime ldt = (LocalDateTime) value;
			return Date.from(ldt.atZone(java.time.ZoneOffset.UTC).toInstant());
		}
		if (value instanceof LocalDate) {
			LocalDate ld = (LocalDate) value;
			return Date.from(ld.atStartOfDay(java.time.ZoneOffset.UTC).toInstant());
		}
		if (value instanceof java.util.Date) {
			return (Date) value;
		}
		if (value instanceof Number) {
			return org.apache.poi.ss.usermodel.DateUtil.getJavaDate(((Number) value).doubleValue());
		}
		if (value instanceof String) {
			String s = ((String) value).trim();
			if (s.isEmpty()) return null;
			try {
				if (s.endsWith("Z") || s.contains("+")) {
					java.time.Instant instant = java.time.Instant.parse(s);
					return Date.from(instant);
				} else if (s.contains("T")) {
					LocalDateTime ldt = LocalDateTime.parse(s, java.time.format.DateTimeFormatter.ISO_LOCAL_DATE_TIME);
					return Date.from(ldt.atZone(java.time.ZoneOffset.UTC).toInstant());
				} else {
					LocalDate ld = LocalDate.parse(s, java.time.format.DateTimeFormatter.ISO_LOCAL_DATE);
					return Date.from(ld.atStartOfDay(java.time.ZoneOffset.UTC).toInstant());
				}
			} catch (Exception e) {
				// Not an ISO date string
			}
		}
		return null;
	}

	@Override
	public CellType getCellType() {
		String formula = range.getFormula();
		if (formula != null && !formula.isEmpty()) {
			return CellType.FORMULA;
		}
		Object val = range.getValue();
		if (val == null) return CellType.BLANK;
		if (val instanceof String) return CellType.STRING;
		if (val instanceof Number || val instanceof com.github.miachm.sods.OfficeCurrency || val instanceof com.github.miachm.sods.OfficePercentage) return CellType.NUMERIC;
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
		range.setFormula(null);
	}

	@Override
	public void setCellFormula(String formula) {
		range.setFormula(formula);
	}

	@Override
	public String getCellFormula() {
		String formula = range.getFormula();
		return formula == null ? "" : formula;
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
		} else {
			setBlank();
		}
	}
	@Override
	public void setCellValue(RichTextString value) {
		if (value != null) {
			setCellValue(value.getString());
		} else {
			setBlank();
		}
	}
	@Override
	public void setCellValue(boolean value) {
		range.setFormula(null);
		range.setValue(value);
	}
	@Override
	public RichTextString getRichStringCellValue() {
		return new OdsRichTextString(getStringCellValue());
	}
	@Override
	public void removeCellComment() {
		range.setAnnotation(null);
	}

	@Override
	public Comment getCellComment() {
		com.github.miachm.sods.OfficeAnnotation annotation = range.getAnnotation();
		if (annotation != null) {
			OdsComment comment = new OdsComment(annotation);
			comment.setAddress(getAddress());
			return comment;
		}
		return null;
	}

	@Override
	public void setCellComment(Comment comment) {
		if (comment instanceof OdsComment) {
			range.setAnnotation(((OdsComment) comment).getAnnotation());
		} else if (comment != null) {
			RichTextString rts = comment.getString();
			String text = rts != null ? rts.getString() : "";
			range.setAnnotation(new com.github.miachm.sods.OfficeAnnotation(text, java.time.LocalDateTime.now()));
		} else {
			range.setAnnotation(null);
		}
	}
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
			return java.time.LocalDateTime.ofInstant(((Date) value).toInstant(), java.time.ZoneOffset.UTC);
		}
		if (value instanceof Number) {
			Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(((Number) value).doubleValue());
			return java.time.LocalDateTime.ofInstant(date.toInstant(), java.time.ZoneOffset.UTC);
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
	public void removeFormula() {
		range.setFormula(null);
	}

	@Override
	public CellType getCachedFormulaResultType() {
		Object val = range.getValue();
		if (val == null) return CellType.BLANK;
		if (val instanceof String) return CellType.STRING;
		if (val instanceof Number || val instanceof com.github.miachm.sods.OfficeCurrency || val instanceof com.github.miachm.sods.OfficePercentage) return CellType.NUMERIC;
		if (val instanceof Boolean) return CellType.BOOLEAN;
		if (val instanceof java.time.LocalDateTime || val instanceof java.time.LocalDate || val instanceof java.util.Date) {
			return CellType.NUMERIC;
		}
		return CellType.STRING;
	}
}
