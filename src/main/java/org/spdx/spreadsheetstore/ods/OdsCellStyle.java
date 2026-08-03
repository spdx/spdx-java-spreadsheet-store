/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

/**
 * Adapter for Apache POI {@link CellStyle} over the SODS {@link com.github.miachm.sods.Style}.
 * Manages borders, fonts, alignments, and backgrounds mapping for ODS cells.
 */
public class OdsCellStyle implements CellStyle {

	private final com.github.miachm.sods.Style style = new com.github.miachm.sods.Style();
	private final com.github.miachm.sods.Borders borders = new com.github.miachm.sods.Borders();
	private boolean wrapText = false;
	private int fontIndex = -1;
	private short fillForegroundColor = 0;
	private FillPatternType fillPattern = FillPatternType.NO_FILL;
	private HorizontalAlignment horizontalAlignment = HorizontalAlignment.GENERAL;
	private VerticalAlignment verticalAlignment = VerticalAlignment.BOTTOM;
	private short dataFormat = 0;
	private short index = 0;

	private BorderStyle borderBottom = BorderStyle.NONE;
	private BorderStyle borderLeft = BorderStyle.NONE;
	private BorderStyle borderRight = BorderStyle.NONE;
	private BorderStyle borderTop = BorderStyle.NONE;

	private final OdsWorkbook workbook;

	/**
	 * Creates an ODS cell style adapter.
	 *
	 * @param workbook Parent ODS workbook.
	 */
	public OdsCellStyle(OdsWorkbook workbook) {
		this.workbook = workbook;
	}

	/**
	 * Returns the underlying SODS Style instance.
	 *
	 * @return The underlying SODS Style.
	 */
	com.github.miachm.sods.Style getSodsStyle() {
		return style;
	}

	@Override
	public void setAlignment(HorizontalAlignment align) {
		this.horizontalAlignment = align;
		if (align == HorizontalAlignment.CENTER || align == HorizontalAlignment.CENTER_SELECTION) {
			style.setTextAligment(com.github.miachm.sods.Style.TEXT_ALIGMENT.Center);
		} else if (align == HorizontalAlignment.RIGHT) {
			style.setTextAligment(com.github.miachm.sods.Style.TEXT_ALIGMENT.Right);
		} else if (align == HorizontalAlignment.LEFT) {
			style.setTextAligment(com.github.miachm.sods.Style.TEXT_ALIGMENT.Left);
		}
	}

	@Override
	public HorizontalAlignment getAlignment() {
		return horizontalAlignment;
	}

	@Override
	public void setVerticalAlignment(VerticalAlignment align) {
		this.verticalAlignment = align;
		if (align == VerticalAlignment.TOP) {
			style.setVerticalTextAligment(com.github.miachm.sods.Style.VERTICAL_TEXT_ALIGMENT.Top);
		} else if (align == VerticalAlignment.CENTER) {
			style.setVerticalTextAligment(com.github.miachm.sods.Style.VERTICAL_TEXT_ALIGMENT.Middle);
		} else if (align == VerticalAlignment.BOTTOM) {
			style.setVerticalTextAligment(com.github.miachm.sods.Style.VERTICAL_TEXT_ALIGMENT.Bottom);
		}
	}

	@Override
	public VerticalAlignment getVerticalAlignment() {
		return verticalAlignment;
	}

	@Override
	public void setBorderBottom(BorderStyle border) {
		this.borderBottom = border;
		updateBorders();
	}

	@Override
	public BorderStyle getBorderBottom() {
		return borderBottom;
	}

	@Override
	public void setBorderLeft(BorderStyle border) {
		this.borderLeft = border;
		updateBorders();
	}

	@Override
	public BorderStyle getBorderLeft() {
		return borderLeft;
	}

	@Override
	public void setBorderRight(BorderStyle border) {
		this.borderRight = border;
		updateBorders();
	}

	@Override
	public BorderStyle getBorderRight() {
		return borderRight;
	}

	@Override
	public void setBorderTop(BorderStyle border) {
		this.borderTop = border;
		updateBorders();
	}

	@Override
	public BorderStyle getBorderTop() {
		return borderTop;
	}

	private void updateBorders() {
		boolean hasBottom = borderBottom != null && borderBottom != BorderStyle.NONE;
		borders.setBorderBottom(hasBottom);
		if (hasBottom) {
			borders.setBorderBottomProperties(getBorderProperties(borderBottom));
		}
		
		boolean hasLeft = borderLeft != null && borderLeft != BorderStyle.NONE;
		borders.setBorderLeft(hasLeft);
		if (hasLeft) {
			borders.setBorderLeftProperties(getBorderProperties(borderLeft));
		}
		
		boolean hasRight = borderRight != null && borderRight != BorderStyle.NONE;
		borders.setBorderRight(hasRight);
		if (hasRight) {
			borders.setBorderRightProperties(getBorderProperties(borderRight));
		}
		
		boolean hasTop = borderTop != null && borderTop != BorderStyle.NONE;
		borders.setBorderTop(hasTop);
		if (hasTop) {
			borders.setBorderTopProperties(getBorderProperties(borderTop));
		}
		
		style.setBorders(borders);
	}

	private String getBorderProperties(BorderStyle borderStyle) {
		if (borderStyle == BorderStyle.NONE) {
			return null;
		}
		switch (borderStyle) {
			case THICK:
				return "0.07cm solid #000000";
			case MEDIUM:
				return "0.05cm solid #000000";
			case DOUBLE:
				return "0.05cm double #000000";
			case DASHED:
			case MEDIUM_DASHED:
				return "0.035cm dashed #000000";
			case DOTTED:
				return "0.035cm dotted #000000";
			default:
				return "0.035cm solid #000000";
		}
	}

	@Override
	public void setFillForegroundColor(short bg) {
		this.fillForegroundColor = bg;
		com.github.miachm.sods.Color color = getSodsColor(bg);
		if (color != null) {
			style.setBackgroundColor(color);
		}
	}

	private com.github.miachm.sods.Color getSodsColor(short indexedColor) {
		java.util.Map<Integer, org.apache.poi.hssf.util.HSSFColor> map = org.apache.poi.hssf.util.HSSFColor.getIndexHash();
		org.apache.poi.hssf.util.HSSFColor hc = map.get((int) indexedColor);
		if (hc != null && hc.getTriplet() != null) {
			short[] rgb = hc.getTriplet();
			return new com.github.miachm.sods.Color(rgb[0], rgb[1], rgb[2]);
		}
		return null;
	}

	@Override
	public short getFillForegroundColor() {
		return fillForegroundColor;
	}

	@Override
	public void setFillPattern(FillPatternType fp) {
		this.fillPattern = fp;
	}

	@Override
	public FillPatternType getFillPattern() {
		return fillPattern;
	}

	@Override
	public void setFont(Font font) {
		if (font == null) return;
		this.fontIndex = font.getIndex();
		style.setFontFamily(font.getFontName());
		style.setFontSize((int) Math.round((double) font.getFontHeight() / 20.0));
		style.setBold(font.getBold());
		style.setItalic(font.getItalic());
		style.setUnderline(font.getUnderline() != org.apache.poi.ss.usermodel.Font.U_NONE);
		style.setLineThrough(font.getStrikeout());
		short fontColorIdx = font.getColor();
		com.github.miachm.sods.Color fontColor = getSodsColor(fontColorIdx);
		if (fontColor != null) {
			style.setFontColor(fontColor);
		}
	}

	/**
	 * Sets the data format for the cell.
	 * <p>
	 * Due to limitations in the underlying SODS library, exact POI data format patterns
	 * are not fully supported. If the provided format index corresponds to a recognized
	 * Date format, it will be mapped to the ISO date style ("YYYY-MM-DD") to ensure
	 * dates are properly displayed. Other formats are currently stored but may not
	 * affect the final ODS cell rendering.
	 * </p>
	 * <p>
	 * This is display-only: the full {@code HH:mm:ss} timestamp is still stored
	 * in {@code office:date-value} and returned unchanged by
	 * {@link Cell#getDateCellValue()}, even though time-inclusive formats
	 * render as date-only in spreadsheet applications.
	 * </p>
	 *
	 * @param fmt the data format index
	 */
	@Override
	public void setDataFormat(short fmt) {
		this.dataFormat = fmt;
		if (workbook != null) {
			String formatStr = workbook.createDataFormat().getFormat(fmt);
			if (DateUtil.isADateFormat(fmt, formatStr)) {
				style.setDataStyle("YYYY-MM-DD");
			}
		}
	}

	@Override
	public short getDataFormat() {
		return dataFormat;
	}

	@Override
	public void setWrapText(boolean wrapped) {
		this.wrapText = wrapped;
		style.setWrap(wrapped);
	}

	@Override
	public boolean getWrapText() {
		return wrapText;
	}

	@Override
	public int getFontIndex() {
		return fontIndex;
	}

	@Override
	public int getFontIndexAsInt() {
		return fontIndex;
	}

	@Override
	public short getIndex() {
		return index;
	}

	@Override
	public void cloneStyleFrom(CellStyle source) {}
	@Override
	public void setBottomBorderColor(short color) {}
	@Override
	public short getBottomBorderColor() { return 0; }
	@Override
	public void setLeftBorderColor(short color) {}
	@Override
	public short getLeftBorderColor() { return 0; }
	@Override
	public void setRightBorderColor(short color) {}
	@Override
	public short getRightBorderColor() { return 0; }
	@Override
	public void setTopBorderColor(short color) {}
	@Override
	public short getTopBorderColor() { return 0; }
	@Override
	public void setFillBackgroundColor(short bg) {}
	@Override
	public short getFillBackgroundColor() { return 0; }
	@Override
	public void setHidden(boolean hidden) {}
	@Override
	public boolean getHidden() { return false; }
	@Override
	public void setLocked(boolean locked) {}
	@Override
	public boolean getLocked() { return false; }
	@Override
	public void setQuotePrefixed(boolean quotePrefix) {}
	@Override
	public boolean getQuotePrefixed() { return false; }
	@Override
	public void setIndention(short indent) {}
	@Override
	public short getIndention() { return 0; }
	@Override
	public void setRotation(short rotation) {}
	@Override
	public short getRotation() { return 0; }
	@Override
	public String getDataFormatString() { return ""; }
	@Override
	public void setShrinkToFit(boolean shrinkToFit) {}
	@Override
	public boolean getShrinkToFit() { return false; }

	@Override
	public void invalidateCachedProperties() {}

	@Override
	public java.util.EnumMap<org.apache.poi.ss.usermodel.CellPropertyType, Object> getFormatProperties() {
		return new java.util.EnumMap<>(org.apache.poi.ss.usermodel.CellPropertyType.class);
	}

	@Override
	public Color getFillForegroundColorColor() { return null; }

	@Override
	public Color getFillBackgroundColorColor() { return null; }

	@Override
	public void setFillForegroundColor(Color color) {}

	@Override
	public void setFillBackgroundColor(Color color) {}
}
