package org.spdx.spreadsheetstore;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultModelStore;
import org.spdx.core.ModelRegistry;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.SpdxModelInfoV2_X;
import org.spdx.library.model.v3_0_0.SpdxModelInfoV3_0;
import org.spdx.storage.IModelStore;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

public class ExtractedLicenseInfoSheetTest extends TestCase {


	private static final String DOCUMENT_URI = "http://spdx.org/test";
	IModelStore modelStore;
	ModelCopyManager copyManager;
	/* (non-Javadoc)
	 * @see junit.framework.TestCase#setUp()
	 */
	protected void setUp() throws Exception {
		super.setUp();
		modelStore = new InMemSpdxStore();
		copyManager = new ModelCopyManager();
		ModelRegistry.getModelRegistry().registerModel(new SpdxModelInfoV2_X());
		ModelRegistry.getModelRegistry().registerModel(new SpdxModelInfoV3_0());
		DefaultModelStore.initialize(modelStore, DOCUMENT_URI, copyManager);
	}

	/* (non-Javadoc)
	 * @see junit.framework.TestCase#tearDown()
	 */
	protected void tearDown() throws Exception {
		super.tearDown();
	}
	
	public void testCreate() {
		Workbook wb = new HSSFWorkbook();
		ExtractedLicenseInfoSheet.create(wb, "Extracted Licenses");
		ExtractedLicenseInfoSheet sheet = ExtractedLicenseInfoSheet.openVersion(wb, "Extracted Licenses", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}
	
	public void testAdd() {
		Workbook wb = new HSSFWorkbook();
		ExtractedLicenseInfoSheet.create(wb, "Extracted Licenses");
		ExtractedLicenseInfoSheet sheet = ExtractedLicenseInfoSheet.openVersion(wb, "Extracted Licenses", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		String id1 = "LicenseRef-1";
		String text1 = "text1";
		String name1 = "name1";
		Collection<String> crossrefs1 = Arrays.asList(new String[] {"http://www.one", "http:www.two"});
		String comment1 = "comment1";
		
		String id2 = "LicenseRef-2";
		String text2 = "text2";
		String name2 = "name2";
		Collection<String> crossrefs2 = Arrays.asList(new String[] {"http://www.three"});
		String comment2 = "comment2";
		
		String id3 = "LicenseRef-3";
		String text3 = "text3";
		String name3 = null;
		Collection<String> crossrefs3 = new ArrayList<>();
		String comment3 = null;
		sheet.add(id1, text1, name1, crossrefs1, comment1);
		sheet.add(id2, text2, name2, crossrefs2, comment2);
		sheet.add(id3, text3, name3, crossrefs3, comment3);
		assertEquals(id1, sheet.getIdentifier(sheet.getFirstDataRow()));
		assertEquals(text1, sheet.getExtractedText(sheet.getFirstDataRow()));
		assertEquals(name1, sheet.getLicenseName(sheet.getFirstDataRow()));
		assertTrue(collectionsSame(crossrefs1, sheet.getCrossRefUrls(sheet.getFirstDataRow())));
		assertEquals(comment1, sheet.getComment(sheet.getFirstDataRow()));
		
		assertEquals(id2, sheet.getIdentifier(sheet.getFirstDataRow()+1));
		assertEquals(text2, sheet.getExtractedText(sheet.getFirstDataRow()+1));
		assertEquals(name2, sheet.getLicenseName(sheet.getFirstDataRow()+1));
		assertTrue(collectionsSame(crossrefs2, sheet.getCrossRefUrls(sheet.getFirstDataRow()+1)));
		assertEquals(comment2, sheet.getComment(sheet.getFirstDataRow()+1));
		
		assertEquals(id3, sheet.getIdentifier(sheet.getFirstDataRow()+2));
		assertEquals(text3, sheet.getExtractedText(sheet.getFirstDataRow()+2));
		assertEquals(name3, sheet.getLicenseName(sheet.getFirstDataRow()+2));
		assertTrue(collectionsSame(crossrefs3, sheet.getCrossRefUrls(sheet.getFirstDataRow()+2)));
		assertEquals(comment3, sheet.getComment(sheet.getFirstDataRow()+2));

	}

	private boolean collectionsSame(Collection<String> c1, Collection<String> c2) {
		if (c1.size() != c2.size()) {
			return false;
		}
		for (String s1:c1) {
			boolean found = false;
			if (c2.contains(s1)) {
				found = true;
				break;
			}
			if (!found) {
				return false;
			}
		}
		return true;
	}

}
