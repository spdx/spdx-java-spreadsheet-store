/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;

/**
 * Adapter for Apache POI {@link RichTextString} for ODS documents.
 */
public class OdsRichTextString implements RichTextString {

	private String text;

	public OdsRichTextString(String text) {
		this.text = text != null ? text : "";
	}

	@Override
	public void applyFont(int startIndex, int endIndex, short fontIndex) {
	}

	@Override
	public void applyFont(int startIndex, int endIndex, Font font) {
	}

	@Override
	public void applyFont(Font font) {
	}

	@Override
	public void clearFormatting() {
	}

	@Override
	public String getString() {
		return text;
	}

	@Override
	public int length() {
		return text.length();
	}

	@Override
	public int numFormattingRuns() {
		return 0;
	}

	@Override
	public int getIndexOfFormattingRun(int index) {
		return 0;
	}

	@Override
	public void applyFont(short fontIndex) {
	}

	@Override
	public String toString() {
		return text;
	}
}
