/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.function.Supplier;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;
import org.spdx.spreadsheetstore.SpreadsheetTestUtils;

import static org.junit.Assert.assertEquals;

/**
 * Row and cell bounds of ODS match XLS and XLSX for the same sequence of calls.
 */
public class RowApiParityTest {

	private static final int PROBE_ROWS = 7;

	private interface Layout {
		void build(Workbook workbook, Sheet sheet);
	}

	/** Bounds and existence of rows 0..PROBE_ROWS-1 and their cells */
	private static String snapshot(Sheet sheet) {
		StringBuilder sb = new StringBuilder();
		sb.append("first=").append(sheet.getFirstRowNum()).append(" last=").append(sheet.getLastRowNum());
		for (int i = 0; i < PROBE_ROWS; i++) {
			Row row = sheet.getRow(i);
			sb.append(" | row").append(i).append('=');
			if (row == null) {
				sb.append("null");
			} else {
				sb.append(row.getFirstCellNum()).append("..").append(row.getLastCellNum());
			}
		}
		return sb.toString();
	}

	/** Builds the layout on a new sheet in XLS, XLSX and ODS, and asserts all snapshots are equal. */
	private static void assertParity(Layout layout) {
		String expected = null;
		for (Supplier<Workbook> factory : SpreadsheetTestUtils.WORKBOOK_FACTORIES) {
			try (Workbook workbook = factory.get()) {
				Sheet sheet = workbook.createSheet("S");
				layout.build(workbook, sheet);
				String actual = snapshot(sheet);
				if (expected == null) {
					expected = actual;
				}
				assertEquals(workbook.getClass().getSimpleName(), expected, actual);
			} catch (java.io.IOException e) {
				throw new AssertionError(e);
			}
		}
	}

	/** Builds the layout, writes and reloads the workbook, applies the change, then asserts as {@link #assertParity}. */
	private static void assertParityAfterReload(Layout layout, Layout change) {
		String expected = null;
		for (Supplier<Workbook> factory : SpreadsheetTestUtils.WORKBOOK_FACTORIES) {
			try (Workbook written = factory.get()) {
				layout.build(written, written.createSheet("S"));
				ByteArrayOutputStream out = new ByteArrayOutputStream();
				written.write(out);
				InputStream in = new ByteArrayInputStream(out.toByteArray());
				try (Workbook workbook = written instanceof OdsWorkbook ? new OdsWorkbook(in) : WorkbookFactory.create(in)) {
					Sheet sheet = workbook.getSheet("S");
					change.build(workbook, sheet);
					String actual = snapshot(sheet);
					if (expected == null) {
						expected = actual;
					}
					assertEquals(workbook.getClass().getSimpleName(), expected, actual);
				}
			} catch (java.io.IOException e) {
				throw new AssertionError(e);
			}
		}
	}

	private static void content(Sheet sheet, int row, int col) {
		Row r = sheet.getRow(row) != null ? sheet.getRow(row) : sheet.createRow(row);
		r.createCell(col).setCellValue(row + ":" + col); // unique: SODS aliases cells of identical rows on load
	}

	@Test
	public void contiguousRows() {
		assertParity((wb, sheet) -> {
			for (int i = 0; i < 3; i++) {
				content(sheet, i, 0);
			}
		});
	}

	@Test
	public void emptySheet() {
		assertParity((wb, sheet) -> { });
	}

	@Test
	public void emptyRowsAfterDataRemoved() {
		assertParity((wb, sheet) -> {
			for (int i = 0; i < 5; i++) {
				sheet.createRow(i);
			}
			content(sheet, 0, 0);
			content(sheet, 1, 0);
			for (int i = 2; i < 5; i++) {
				sheet.removeRow(sheet.getRow(i));
			}
		});
	}

	@Test
	public void emptyRowsBeforeDataRemoved() {
		assertParity((wb, sheet) -> {
			for (int i = 0; i < 4; i++) {
				sheet.createRow(i);
			}
			content(sheet, 2, 0);
			content(sheet, 3, 0);
			sheet.removeRow(sheet.getRow(0));
			sheet.removeRow(sheet.getRow(1));
		});
	}

	@Test
	public void firstAndLastRowRemoved() {
		assertParity((wb, sheet) -> {
			for (int i = 0; i < 4; i++) {
				content(sheet, i, 0);
			}
			sheet.removeRow(sheet.getRow(0));
			sheet.removeRow(sheet.getRow(3));
		});
	}

	@Test
	public void allRowsRemoved() {
		assertParity((wb, sheet) -> {
			content(sheet, 0, 0);
			content(sheet, 1, 0);
			sheet.removeRow(sheet.getRow(1));
			sheet.removeRow(sheet.getRow(0));
		});
	}

	@Test
	public void loadedLastRowRemoved() {
		assertParityAfterReload(
				(wb, sheet) -> {
					for (int i = 0; i < 4; i++) {
						content(sheet, i, 0);
					}
				},
				(wb, sheet) -> {
					Row last = sheet.getRow(3);
					last.getCell(0); // OdsRow.clear only clears cells already read
					sheet.removeRow(last);
					sheet.getRow(3); // probe before bounds are read
				});
	}

	@Test
	public void cellBounds() {
		assertParity((wb, sheet) -> {
			content(sheet, 0, 1);
			content(sheet, 0, 2);
			content(sheet, 1, 5);
			sheet.createRow(2);
			sheet.getRow(2).createCell(3); // blank cell
			CellStyle style = wb.createCellStyle();
			style.setAlignment(HorizontalAlignment.CENTER);
			sheet.createRow(3).createCell(4).setCellStyle(style); // styled blank cell
		});
	}
}
