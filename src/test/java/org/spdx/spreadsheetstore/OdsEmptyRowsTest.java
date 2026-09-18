/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.TreeSet;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxModelFactory;
import org.spdx.library.model.v2.SpdxConstantsCompatV2;
import org.spdx.library.model.v2.SpdxDocument;
import org.spdx.library.model.v2.SpdxElement;
import org.spdx.spreadsheetstore.ods.OdsWorkbook;
import org.spdx.storage.simple.InMemSpdxStore;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;


/**
 * ODS files from LibreOffice end every sheet with empty rows.
 * Those rows must be ignored: content matches the XLSX example.
 * Fixtures: <code>TestFiles/ods</code>, from <code>TestFiles/generate_ods_fixtures.py</code>.
 */
public class OdsEmptyRowsTest {

	private static final String TEST_FILES = "TestFiles";
	private static final String ODS_DIR = TEST_FILES + File.separator + "ods";
	private static final String V2_3 = "SPDXSpreadsheetExample-v2.3";
	private static final String[] EXAMPLES = new String[] {"SPDXSpreadsheetExample-2.0", "SPDXSpreadsheetExample-v2.2", V2_3};
	/** v2.3 variants with extra empty rows that must load the same as the XLSX */
	private static final String[] EMPTY_ROW_VARIANTS = new String[] {
			V2_3 + "-trailing-repeated-5000", V2_3 + "-trailing-styled-blank",
			V2_3 + "-trailing-many-columns", V2_3 + "-leading-empty-row"};

	/** Element IDs and counts found in a spreadsheet */
	private static class Contents {
		List<String> files = new ArrayList<>();
		List<String> snippets = new ArrayList<>();
		List<String> packages = new ArrayList<>();
		int relationships = 0;
		int annotations = 0;
	}

	private Contents load(String path) throws InvalidSPDXAnalysisException, IOException {
		SpreadsheetStore sst = new SpreadsheetStore(new InMemSpdxStore());
		try (FileInputStream stream = new FileInputStream(path)) {
			sst.deSerialize(stream, false);
		}
		ModelCopyManager cm = new ModelCopyManager();
		String documentUri;
		try (Stream<?> docs = SpdxModelFactory.getSpdxObjects(sst, cm,
				SpdxConstantsCompatV2.CLASS_SPDX_DOCUMENT, null, null)) {
			List<?> allDocs = docs.collect(Collectors.toList());
			assertEquals(1, allDocs.size());
			documentUri = ((SpdxDocument)allDocs.get(0)).getDocumentUri();
		}
		Contents retval = new Contents();
		SpdxDocument doc = new SpdxDocument(sst, documentUri, cm, false);
		retval.annotations = doc.getAnnotations().size();
		retval.relationships = doc.getRelationships().size();
		retval.files = ids(sst, cm, SpdxConstantsCompatV2.CLASS_SPDX_FILE, documentUri);
		retval.snippets = ids(sst, cm, SpdxConstantsCompatV2.CLASS_SPDX_SNIPPET, documentUri);
		retval.packages = ids(sst, cm, SpdxConstantsCompatV2.CLASS_SPDX_PACKAGE, documentUri);
		return retval;
	}

	private List<String> ids(SpreadsheetStore sst, ModelCopyManager cm, String type, String documentUri)
			throws InvalidSPDXAnalysisException {
		List<String> retval = new ArrayList<>();
		try (Stream<?> elements = SpdxModelFactory.getSpdxObjects(sst, cm, type, documentUri, documentUri + "#")) {
			elements.forEach(e -> retval.add(((SpdxElement)e).getId()));
		}
		Collections.sort(retval);
		return retval;
	}

	private void assertSameContents(Contents expected, Contents actual) {
		assertEquals(expected.files, actual.files);
		assertEquals(expected.snippets, actual.snippets);
		assertEquals(expected.packages, actual.packages);
		assertEquals(expected.relationships, actual.relationships);
		assertEquals(expected.annotations, actual.annotations);
	}

	private void assertNoAnonymousSnippets(Contents contents) {
		for (String id : contents.snippets) {
			assertFalse("Phantom snippet from an empty row: " + id, id.contains("__anon__"));
		}
		assertEquals(contents.snippets.size(), new TreeSet<>(contents.snippets).size());
	}

	@Test
	public void libreOfficeConvertedOdsMatchesXlsx() throws InvalidSPDXAnalysisException, IOException {
		for (String example : EXAMPLES) {
			Contents expected = load(TEST_FILES + File.separator + example + ".xlsx");
			Contents ods = load(ODS_DIR + File.separator + example + ".ods");
			assertNoAnonymousSnippets(ods);
			assertSameContents(expected, ods);
		}
	}

	@Test
	public void emptyRowsAreIgnored() throws InvalidSPDXAnalysisException, IOException {
		Contents expected = load(TEST_FILES + File.separator + V2_3 + ".xlsx");
		assertFalse(expected.snippets.isEmpty());
		for (String variant : EMPTY_ROW_VARIANTS) {
			Contents ods = load(ODS_DIR + File.separator + variant + ".ods");
			assertNoAnonymousSnippets(ods);
			assertSameContents(expected, ods);
		}
	}

	@Test
	public void headerOnlySnippetSheetHasNoSnippets() throws InvalidSPDXAnalysisException, IOException {
		Contents expected = load(TEST_FILES + File.separator + V2_3 + ".xlsx");
		Contents ods = load(ODS_DIR + File.separator + V2_3 + "-snippets-header-only.ods");
		assertTrue(ods.snippets.isEmpty());
		assertEquals(expected.files, ods.files);
		assertEquals(expected.packages, ods.packages);
		assertEquals(expected.annotations, ods.annotations);
	}

	/** Row bounds of every sheet equal the XLSX's, shifted by the rows added before the data. */
	private void assertRowBounds(String odsPath, int rowOffset) throws IOException {
		try (FileInputStream xlsxStream = new FileInputStream(TEST_FILES + File.separator + V2_3 + ".xlsx");
				FileInputStream odsStream = new FileInputStream(odsPath);
				Workbook xlsx = WorkbookFactory.create(xlsxStream);
				Workbook ods = new OdsWorkbook(odsStream)) {
			assertEquals(xlsx.getNumberOfSheets(), ods.getNumberOfSheets());
			for (int i = 0; i < xlsx.getNumberOfSheets(); i++) {
				Sheet expected = xlsx.getSheetAt(i);
				Sheet actual = ods.getSheet(expected.getSheetName());
				String name = expected.getSheetName();
				assertEquals(name, expected.getFirstRowNum() + rowOffset, actual.getFirstRowNum());
				assertEquals(name, expected.getLastRowNum() + rowOffset, actual.getLastRowNum());
				if (rowOffset > 0) {
					assertNull(name, actual.getRow(0));
				}
				assertNull(name, actual.getRow(actual.getLastRowNum() + 1));
				Row header = actual.getRow(actual.getFirstRowNum());
				assertEquals(name, expected.getRow(expected.getFirstRowNum()).getFirstCellNum(), header.getFirstCellNum());
				assertEquals(name, expected.getRow(expected.getFirstRowNum()).getLastCellNum(), header.getLastCellNum());
			}
		}
	}

	@Test
	public void trailingEmptyRowsAreNotRows() throws IOException {
		assertRowBounds(ODS_DIR + File.separator + V2_3 + ".ods", 0);
	}

	@Test
	public void leadingEmptyRowIsNotARow() throws IOException {
		assertRowBounds(ODS_DIR + File.separator + V2_3 + "-leading-empty-row.ods", 1);
	}
}
