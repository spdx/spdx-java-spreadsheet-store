/**
 * Copyright (c) 2020 Source Auditor Inc.
 *
 * SPDX-License-Identifier: Apache-2.0
 * 
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 *
 *       http://www.apache.org/licenses/LICENSE-2.0
 *
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */
package org.spdx.spreadsheetstore;

import java.util.Arrays;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.Checksum;
import org.spdx.library.model.ExternalDocumentRef;
import org.spdx.library.model.ModelObject;
import org.spdx.library.model.SpdxDocument;
import org.spdx.library.model.SpdxModelFactory;
import org.spdx.library.model.enumerations.ChecksumAlgorithm;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

/**
 * @author gary
 *
 */
public class DocumentInfoSheetTest extends TestCase {

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
	}

	/* (non-Javadoc)
	 * @see junit.framework.TestCase#tearDown()
	 */
	protected void tearDown() throws Exception {
		super.tearDown();
	}
	
	public void testCreate() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		String result = originsSheet.verify();
		if (result != null && !result.isEmpty()) {
			fail(result);
		}
	}

	/**
	 * Test method for {@link org.spdx.SpdxSpreadsheet.DocumentInfoSheet#setSPDXVersion(java.lang.String)}.
	 * @throws SpreadsheetException 
	 */
	
	public void testSetSPDXVersion() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		String spdxVersion = "2.0";
		originsSheet.setSPDXVersion(spdxVersion);
		assertEquals(spdxVersion, originsSheet.getSPDXVersion());
		spdxVersion = "2.2";
		originsSheet.setSPDXVersion(spdxVersion);
		assertEquals(spdxVersion, originsSheet.getSPDXVersion());
	}

	/**
	 * Test method for {@link org.spdx.SpdxSpreadsheet.DocumentInfoSheet#setCreatedBy(java.lang.String[])}.
	 */
	
	public void testSetCreatedBy() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		Collection<String> createdBys = Arrays.asList(new String[] {"Person: Gary O'Neall", "Tool: Source Auditor Scanner"});
		originsSheet.setCreatedBy(createdBys);
		compareStrings(createdBys, originsSheet.getCreatedBy());
		createdBys = Arrays.asList(new String[] {"Tool: FOSSOlogy"});
		originsSheet.setCreatedBy(createdBys);
		compareStrings(createdBys, originsSheet.getCreatedBy());
	}

	/**
	 * @param s1
	 * @param s2
	 */
	private void compareStrings(Collection<String> s1, Collection<String> s2) {
		assertEquals(s1.size(), s2.size());
		s1.forEach(s -> {
			if (!s2.contains(s)) {
				fail("Strings different.  Missing "+s);
			}
		});
	}

	public void testSetDataLicense() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		String licenseId = "CC0";
		originsSheet.setDataLicense(licenseId);
		assertEquals(licenseId, originsSheet.getDataLicense());
		licenseId = "GPL-2.0+";
		originsSheet.setDataLicense(licenseId);
		assertEquals(licenseId, originsSheet.getDataLicense());
	}

	public void testSetAuthorComments() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		String comment = "comment1";
		originsSheet.setAuthorComments(comment);
		assertEquals(comment, originsSheet.getAuthorComments());
		comment = "comment which is different";
		originsSheet.setAuthorComments(comment);
		assertEquals(comment, originsSheet.getAuthorComments());
	}

	public void testSetCreated() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		Date created = new Date();
		originsSheet.setCreated(created);
		assertEquals(created.toString(), originsSheet.getCreated().toString());
	}

	public void testGetDocumentomment() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		String comment = "comment1";
		originsSheet.setDocumentComment(comment);
		assertEquals(comment, originsSheet.getDocumentComment());
		comment = "comment which is different";
		originsSheet.setDocumentComment(comment);
		assertEquals(comment, originsSheet.getDocumentComment());
	}

	public void testSetLicenseListVersion() throws SpreadsheetException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		String ver = "1.19";
		originsSheet.setLicenseListVersion(ver);
		assertEquals(ver, originsSheet.getLicenseListVersion());
		ver = "1.20";
		originsSheet.setLicenseListVersion(ver);
		assertEquals(ver, originsSheet.getLicenseListVersion());
	}
	
	public void testSetExternalDocumentRefs() throws InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		DocumentInfoSheet.create(wb, "Origins", DOCUMENT_URI);
		DocumentInfoSheet originsSheet = DocumentInfoSheet.openVersion(wb, "Origins", SpdxSpreadsheet.CURRENT_VERSION, modelStore, copyManager);
		SpdxDocument doc = SpdxModelFactory.createSpdxDocument(modelStore, DOCUMENT_URI, new ModelCopyManager());
		String externalDocumentId1 = modelStore.getNextId(IdType.DocumentRef, DOCUMENT_URI);
		String externalDocumentId2 = modelStore.getNextId(IdType.DocumentRef, DOCUMENT_URI);
		String externalDocumentId3 = modelStore.getNextId(IdType.DocumentRef, DOCUMENT_URI);
		String externalDocumentUri1 = "http://external1";
		String externalDocumentUri2 = "http://external2";
		String externalDocumentUri3 = "http://external3";
		Checksum checksum1 = doc.createChecksum(ChecksumAlgorithm.SHA1, "5e85a37701a7cfbd07a41b9e7b20b688bc090ede");
		Checksum checksum2 = doc.createChecksum(ChecksumAlgorithm.MD5, "afaea6bb0ae5b7fdb1d7d09d4268c9ba");
		Checksum checksum3 = doc.createChecksum(ChecksumAlgorithm.SHA1, "7e85a37701a7cfbd07a41b9e7b20b688bc090ede");
		ExternalDocumentRef ref1 = doc.createExternalDocumentRef(externalDocumentId1, externalDocumentUri1, 
				checksum1);
		ExternalDocumentRef ref2 = doc.createExternalDocumentRef(externalDocumentId2, externalDocumentUri2, 
				checksum2);
		ExternalDocumentRef ref3 = doc.createExternalDocumentRef(externalDocumentId3, externalDocumentUri3, 
				checksum3);
		Collection<ExternalDocumentRef> externalDocRefs = Arrays.asList(new ExternalDocumentRef[]{ref1, ref2, ref3});
		originsSheet.setExternalDocumentRefs(externalDocRefs);
		Collection<ExternalDocumentRef> result = originsSheet.getExternalDocumentRefs();
		assertModelObjectsEquiv(externalDocRefs, result);
	}

	private void assertModelObjectsEquiv(Collection<? extends ModelObject> mo1,
			Collection<? extends ModelObject> mo2) {
		assertEquals(mo1.size(), mo2.size());
		mo1.forEach(mo -> {
			boolean found = false;
			Iterator<? extends ModelObject> iter = mo2.iterator();
			while (iter.hasNext() && !found) {
				try {
					if (iter.next().equivalent(mo)) {
						found = true;
					}
				} catch (InvalidSPDXAnalysisException e) {
					fail("Exception comparing: "+e.getMessage());
				}
			}
			if (!found) {
				fail("Compare failed - could not find equiv. to ID "+mo.getId());
			}
		});
	}
}
