/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import org.apache.poi.ss.usermodel.Font;

/**
 * Adapter for Apache POI {@link Font} over SODS styles.
 * Holds font configurations (size, bold, italic, underline, strikeout) for an ODS cell.
 */
public class OdsFont implements Font {

	private final short index;
	private String name = "Arial";
	private short height = 200;
	private boolean bold = false;
	private short color = 0;
	private boolean italic = false;
	private boolean strikeout = false;
	private short typeOffset = 0;
	private byte underline = 0;

	public OdsFont(short index) {
		this.index = index;
	}

	@Override
	public void setFontName(String name) {
		this.name = name;
	}

	@Override
	public String getFontName() {
		return this.name;
	}

	@Override
	public void setFontHeight(short height) {
		this.height = height;
	}

	@Override
	public void setFontHeightInPoints(short height) {
		this.height = (short) (height * 20);
	}

	@Override
	public short getFontHeight() {
		return this.height;
	}

	@Override
	public short getFontHeightInPoints() {
		return (short) (this.height / 20);
	}

	@Override
	public void setItalic(boolean italic) {
		this.italic = italic;
	}

	@Override
	public boolean getItalic() {
		return this.italic;
	}

	@Override
	public void setStrikeout(boolean strikeout) {
		this.strikeout = strikeout;
	}

	@Override
	public boolean getStrikeout() {
		return this.strikeout;
	}

	@Override
	public void setColor(short color) {
		this.color = color;
	}

	@Override
	public short getColor() {
		return this.color;
	}

	@Override
	public void setTypeOffset(short offset) {
		this.typeOffset = offset;
	}

	@Override
	public short getTypeOffset() {
		return this.typeOffset;
	}

	@Override
	public void setUnderline(byte underline) {
		this.underline = underline;
	}

	@Override
	public byte getUnderline() {
		return this.underline;
	}

	@Override
	public int getCharSet() {
		return 0;
	}

	@Override
	public void setCharSet(byte charset) {
	}

	@Override
	public void setCharSet(int charset) {
	}

	@Override
	public void setBold(boolean bold) {
		this.bold = bold;
	}

	@Override
	public boolean getBold() {
		return this.bold;
	}

	@Override
	public int getIndex() {
		return this.index;
	}

	@Override
	public int getIndexAsInt() {
		return this.index;
	}
}
