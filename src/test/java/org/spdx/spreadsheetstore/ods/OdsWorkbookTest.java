/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: Copyright (c) 2026 Source Auditor Inc.
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

/**
 * Unit tests for OdsWorkbook, OdsSheet, OdsRow, and OdsCell implementations.
 */
public class OdsWorkbookTest {

	@Test
	public void testCreateSheetAndRowsAndCells() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("TestSheet");
		assertNotNull(sheet);
		assertEquals("TestSheet", sheet.getSheetName());

		Row row = sheet.createRow(5);
		assertNotNull(row);
		assertEquals(5, row.getRowNum());

		Cell cell = row.createCell(10);
		assertNotNull(cell);
		assertEquals(10, cell.getColumnIndex());
	}

	@Test
	public void testCloneSheet() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet1 = workbook.createSheet("Original");
		Row row = sheet1.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("Hello World");

		Sheet sheet2 = workbook.cloneSheet(0);
		assertNotNull(sheet2);
		assertEquals(2, workbook.getNumberOfSheets());
	}

	@Test
	public void testSheetHiddenState() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("Sheet1");
		assertFalse(workbook.isSheetHidden(0));

		workbook.setSheetHidden(0, true);
		assertTrue(workbook.isSheetHidden(0));

		workbook.setSheetHidden(0, false);
		assertFalse(workbook.isSheetHidden(0));
	}

	@Test
	public void testMergedRegions() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("MergedSheet");
		sheet.createRow(0).createCell(0);
		sheet.createRow(2).createCell(3);

		CellRangeAddress region = new CellRangeAddress(0, 2, 0, 3);
		sheet.addMergedRegion(region);

		assertEquals(1, sheet.getNumMergedRegions());
		CellRangeAddress retrieved = sheet.getMergedRegion(0);
		assertNotNull(retrieved);
		assertEquals(0, retrieved.getFirstRow());
		assertEquals(2, retrieved.getLastRow());
		assertEquals(0, retrieved.getFirstColumn());
		assertEquals(3, retrieved.getLastColumn());
	}

	@Test
	public void testColumnAndRowDimensions() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("DimSheet");
		Row row = sheet.createRow(0);

		sheet.setColumnWidth(2, 5000);
		assertTrue(sheet.getColumnWidth(2) > 0);

		row.setHeightInPoints(25.0f);
		assertTrue(row.getHeightInPoints() > 0);
	}

	@Test
	public void testColumnAndRowVisibility() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("VisSheet");
		Row row = sheet.createRow(1);

		assertFalse(sheet.isColumnHidden(1));
		sheet.setColumnHidden(1, true);
		assertTrue(sheet.isColumnHidden(1));

		assertFalse(row.getZeroHeight());
		row.setZeroHeight(true);
		assertTrue(row.getZeroHeight());
	}

	@Test
	public void testFormulas() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("FormulaSheet");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		cell.setCellFormula("SUM(A1:A5)");
		assertEquals("SUM(A1:A5)", cell.getCellFormula());
		assertEquals(CellType.FORMULA, cell.getCellType());
	}

	@Test
	public void testFreezePanesAndProtection() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("ProtectSheet");

		sheet.createFreezePane(1, 2);
		assertFalse(sheet.getProtect());

		sheet.protectSheet("secret");
		assertTrue(sheet.getProtect());
	}

	@Test
	public void testSheetOrdering() {
		OdsWorkbook workbook = new OdsWorkbook();
		workbook.createSheet("First");
		workbook.createSheet("Second");

		assertEquals(0, workbook.getSheetIndex("First"));
		assertEquals(1, workbook.getSheetIndex("Second"));

		workbook.setSheetOrder("Second", 0);
		assertEquals(0, workbook.getSheetIndex("Second"));
		assertEquals(1, workbook.getSheetIndex("First"));
	}

	@Test
	public void testCellComments() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("CommentSheet");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		OdsComment comment = new OdsComment("This is a comment");
		cell.setCellComment(comment);

		assertNotNull(cell.getCellComment());
		assertEquals("This is a comment", cell.getCellComment().getString().getString());

		cell.removeCellComment();
		org.junit.Assert.assertNull(cell.getCellComment());
	}

	@Test
	public void testCellIteratorNoNulls() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("SparseSheet");
		Row row = sheet.createRow(0);
		row.createCell(0).setCellValue("A");
		row.createCell(5).setCellValue("F"); // Gap columns 1-4

		int cellCount = 0;
		for (Cell cell : row) {
			assertNotNull("Cell iterator must not return null elements", cell);
			cellCount++;
		}
		assertEquals(2, cellCount);
	}

	@Test
	public void testRowIteratorEmptySheet() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("EmptySheet");
		assertEquals(-1, sheet.getFirstRowNum());
		assertEquals(-1, sheet.getLastRowNum());

		int rowCount = 0;
		for (Row row : sheet) {
			rowCount++;
		}
		assertEquals(0, rowCount);
		assertEquals(-1, sheet.getFirstRowNum()); // No synthetic row created
	}

	@Test
	public void testNullBorderHandling() {
		OdsWorkbook workbook = new OdsWorkbook();
		OdsCellStyle style = (OdsCellStyle) workbook.createCellStyle();
		style.setBorderBottom(null);
		org.junit.Assert.assertNull(style.getBorderBottom());
	}

	@Test
	public void testRemoveFormulaAndClear() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("FormulaClear");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);

		cell.setCellFormula("SUM(A1:A5)");
		assertEquals(CellType.FORMULA, cell.getCellType());

		cell.removeFormula();
		assertFalse(cell.getCellType() == CellType.FORMULA);

		cell.setCellFormula("A1+A2");
		cell.setBlank();
		assertEquals(CellType.BLANK, cell.getCellType());
	}

	@Test(expected = IllegalArgumentException.class)
	public void testOutOfBoundsSheetIndexException() {
		OdsWorkbook workbook = new OdsWorkbook();
		workbook.getSheetAt(99);
	}

	@Test
	public void testBuiltinDataFormatFallback() {
		OdsWorkbook workbook = new OdsWorkbook();
		org.apache.poi.ss.usermodel.DataFormat df = workbook.createDataFormat();
		assertEquals("m/d/yy", df.getFormat((short) 14));
	}

	@Test
	public void testIsoDateStringWithZ() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("Dates");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue("2026-08-02T08:34:56Z");
		java.util.Date date = cell.getDateCellValue();
		org.junit.Assert.assertNotNull(date);
	}

	@Test
	public void testHeaderFooterAndPrintSetupStubs() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("Stubs");
		org.junit.Assert.assertNotNull(sheet.getHeader());
		org.junit.Assert.assertNotNull(sheet.getFooter());
		org.junit.Assert.assertNotNull(sheet.getPrintSetup());
		sheet.protectSheet(null);
		org.junit.Assert.assertTrue(sheet.getProtect());
	}

	@Test
	public void testDoubleRowHeightPrecision() {
		OdsWorkbook workbook = new OdsWorkbook();
		Sheet sheet = workbook.createSheet("Heights");
		Row row = sheet.createRow(0);
		row.setHeightInPoints(20.5f);
		assertEquals(20.5f, row.getHeightInPoints(), 0.1f);
	}

	@Test
	public void testFindFontNullName() {
		OdsWorkbook workbook = new OdsWorkbook();
		org.apache.poi.ss.usermodel.Font font = workbook.findFont(false, (short) 0, (short) 200, null, false, false, (short) 0, (byte) 0);
		// Should return without NPE
	}
}
