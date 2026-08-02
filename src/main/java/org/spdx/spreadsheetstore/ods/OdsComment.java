/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.time.LocalDateTime;
import com.github.miachm.sods.OfficeAnnotation;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;

/**
 * Adapter for Apache POI {@link Comment} over SODS {@link OfficeAnnotation}.
 */
public class OdsComment implements Comment {

	private OfficeAnnotation annotation;
	private String author = "";
	private CellAddress address;

	/**
	 * Creates an ODS comment wrapping an existing SODS OfficeAnnotation.
	 *
	 * @param annotation Underlying SODS OfficeAnnotation instance.
	 */
	public OdsComment(OfficeAnnotation annotation) {
		this.annotation = annotation;
	}

	/**
	 * Creates an ODS comment with specified text content.
	 *
	 * @param text Comment text.
	 */
	public OdsComment(String text) {
		this.annotation = new OfficeAnnotation(text, LocalDateTime.now());
	}

	/**
	 * Returns the underlying SODS OfficeAnnotation.
	 *
	 * @return The underlying SODS OfficeAnnotation.
	 */
	OfficeAnnotation getAnnotation() {
		return annotation;
	}

	@Override
	public void setVisible(boolean visible) {
	}

	@Override
	public boolean isVisible() {
		return true;
	}

	@Override
	public CellAddress getAddress() {
		return address;
	}

	@Override
	public void setAddress(CellAddress address) {
		this.address = address;
	}

	@Override
	public void setAddress(int row, int col) {
		this.address = new CellAddress(row, col);
	}

	@Override
	public int getRow() {
		return address != null ? address.getRow() : 0;
	}

	@Override
	public void setRow(int row) {
		int col = getColumn();
		this.address = new CellAddress(row, col);
	}

	@Override
	public int getColumn() {
		return address != null ? address.getColumn() : 0;
	}

	@Override
	public void setColumn(int col) {
		int row = getRow();
		this.address = new CellAddress(row, col);
	}

	@Override
	public String getAuthor() {
		return author;
	}

	@Override
	public void setAuthor(String author) {
		this.author = author;
	}

	@Override
	public RichTextString getString() {
		return new OdsRichTextString(annotation != null ? annotation.getMsg() : "");
	}

	@Override
	public void setString(RichTextString string) {
		String text = string != null ? string.getString() : "";
		this.annotation = new OfficeAnnotation(text, LocalDateTime.now());
	}

	public void setString(String string) {
		this.annotation = new OfficeAnnotation(string != null ? string : "", LocalDateTime.now());
	}

	@Override
	public ClientAnchor getClientAnchor() {
		return null;
	}
}
