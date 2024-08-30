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
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultModelStore;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.core.ModelRegistry;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.ExternalRef;
import org.spdx.library.model.v2.ReferenceType;
import org.spdx.library.model.v2.SpdxModelInfoV2_X;
import org.spdx.library.model.v2.enumerations.ReferenceCategory;
import org.spdx.library.model.v3_0_1.SpdxModelInfoV3_0;
import org.spdx.library.referencetype.ListedReferenceTypes;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

/**
 * @author gary
 *
 */
public class ExternalRefsSheetTest extends TestCase {

	static final String DOCUMENT_NAMSPACE = "http://local/document/namespace";
	
	static final String LOCAL_REFERENCE_TYPE_NAME = "localType";
	static final String FULL_REFRENCE_TYPE_URI = "http://this/is/not/in/the/document#here";
	
	static final String PKG1_ID = "SPDXRef-pkg1";
	static final String PKG2_ID = "SPDXRef-pkg2";
	
	static final String CPE32_NAME = "cpe23Type";
	static final String MAVEN_NAME = "maven-central";
	
	ReferenceType REFERENCE_TYPE_CPE32;
	ReferenceType REFERENCE_TYPE_MAVEN;
	ReferenceType REFERENCE_TYPE_LOCAL_TO_PACKAGE;
	ReferenceType REFERENCE_TYPE_FULL_URI;
	
	ExternalRef EXTERNAL_PKG1_REF1;
	ExternalRef EXTERNAL_PKG1_REF2;
	ExternalRef EXTERNAL_PKG1_REF3;
	ExternalRef EXTERNAL_PKG2_REF1;
	ExternalRef EXTERNAL_PKG2_REF2;
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
		DefaultModelStore.initialize(modelStore, DOCUMENT_NAMSPACE, copyManager);
		REFERENCE_TYPE_CPE32 = ListedReferenceTypes.getListedReferenceTypes().getListedReferenceTypeByName(CPE32_NAME);
		REFERENCE_TYPE_MAVEN = ListedReferenceTypes.getListedReferenceTypes().getListedReferenceTypeByName(MAVEN_NAME);
		REFERENCE_TYPE_FULL_URI = new ReferenceType(FULL_REFRENCE_TYPE_URI);
		REFERENCE_TYPE_LOCAL_TO_PACKAGE = new ReferenceType(DOCUMENT_NAMSPACE + "#" + LOCAL_REFERENCE_TYPE_NAME);
		
		EXTERNAL_PKG1_REF1 = new ExternalRef(modelStore, DOCUMENT_NAMSPACE,
				modelStore.getNextId(IdType.Anonymous), copyManager, true);
		EXTERNAL_PKG1_REF1.setReferenceCategory(ReferenceCategory.SECURITY);
		EXTERNAL_PKG1_REF1.setReferenceLocator("LocatorPkg1Ref1");
		EXTERNAL_PKG1_REF1.setReferenceType(REFERENCE_TYPE_CPE32);
		EXTERNAL_PKG1_REF1.setComment("CommentPkg1Ref1");
		
		EXTERNAL_PKG1_REF2 = new ExternalRef(modelStore, DOCUMENT_NAMSPACE,
				modelStore.getNextId(IdType.Anonymous), copyManager, true);
		EXTERNAL_PKG1_REF2.setReferenceCategory(ReferenceCategory.PACKAGE_MANAGER);
		EXTERNAL_PKG1_REF2.setReferenceLocator("LocatorPkg1Ref2");
		EXTERNAL_PKG1_REF2.setReferenceType(REFERENCE_TYPE_MAVEN);
		EXTERNAL_PKG1_REF2.setComment("CommentPkg1Ref2");		

		EXTERNAL_PKG1_REF3 = new ExternalRef(modelStore, DOCUMENT_NAMSPACE,
				modelStore.getNextId(IdType.Anonymous), copyManager, true);
		EXTERNAL_PKG1_REF3.setReferenceCategory(ReferenceCategory.OTHER);
		EXTERNAL_PKG1_REF3.setReferenceLocator("LocatorPkg1Ref2");
		EXTERNAL_PKG1_REF3.setReferenceType(REFERENCE_TYPE_LOCAL_TO_PACKAGE);
		EXTERNAL_PKG1_REF3.setComment("CommentPkg1Ref2");		

		EXTERNAL_PKG2_REF1 = new ExternalRef(modelStore, DOCUMENT_NAMSPACE,
				modelStore.getNextId(IdType.Anonymous), copyManager, true);
		EXTERNAL_PKG2_REF1.setReferenceCategory(ReferenceCategory.SECURITY);
		EXTERNAL_PKG2_REF1.setReferenceLocator("LocatorPkg2Ref1");
		EXTERNAL_PKG2_REF1.setReferenceType(REFERENCE_TYPE_CPE32);
		EXTERNAL_PKG2_REF1.setComment("CommentPk21Ref1");		

		EXTERNAL_PKG2_REF2 = new ExternalRef(modelStore, DOCUMENT_NAMSPACE,
				modelStore.getNextId(IdType.Anonymous), copyManager, true);
		EXTERNAL_PKG2_REF2.setReferenceCategory(ReferenceCategory.OTHER);
		EXTERNAL_PKG2_REF2.setReferenceLocator("LocatorPkg2Ref2");
		EXTERNAL_PKG2_REF2.setReferenceType(REFERENCE_TYPE_FULL_URI);
		EXTERNAL_PKG2_REF2.setComment("CommentPkg2Ref2");				
	}

	/* (non-Javadoc)
	 * @see junit.framework.TestCase#tearDown()
	 */
	protected void tearDown() throws Exception {
		super.tearDown();
	}
	
	public void testCreate() {
		Workbook wb = new HSSFWorkbook();
		ExternalRefsSheet.create(wb, "External References");
		ExternalRefsSheet sheet = new ExternalRefsSheet(wb, "External References", 
				modelStore, DOCUMENT_NAMSPACE, copyManager);
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}
	public void testAddGet() throws InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		ExternalRefsSheet.create(wb, "External Refs");
		ExternalRefsSheet externalRefsSheet = new ExternalRefsSheet(wb, "External Refs", modelStore, 
				DOCUMENT_NAMSPACE, copyManager);
		externalRefsSheet.add(PKG1_ID, EXTERNAL_PKG1_REF1);
		externalRefsSheet.add(PKG2_ID, EXTERNAL_PKG2_REF1);
		externalRefsSheet.add(PKG1_ID, EXTERNAL_PKG1_REF2);
		externalRefsSheet.add(PKG1_ID, EXTERNAL_PKG1_REF3);
		externalRefsSheet.add(PKG2_ID, EXTERNAL_PKG2_REF2);
		assertTrue(externalRefsSheet.verify() == null);
		
		List<ExternalRef> expectedPkg1 = Arrays.asList(new ExternalRef[] {EXTERNAL_PKG1_REF1,
				EXTERNAL_PKG1_REF2, EXTERNAL_PKG1_REF3});
		List<ExternalRef> expectedPkg2 = Arrays.asList(new ExternalRef[] {EXTERNAL_PKG2_REF1,
				EXTERNAL_PKG2_REF2});
		
		List<ExternalRef> result = externalRefsSheet.getExternalRefsForPkgid(PKG1_ID);
		assertTrue(isListsEquivalent(expectedPkg1, result));
		
		result = externalRefsSheet.getExternalRefsForPkgid(PKG2_ID);
		assertTrue(isListsEquivalent(expectedPkg2, result));
	}

	private boolean isListsEquivalent(List<ExternalRef> l1, List<ExternalRef> l2) throws InvalidSPDXAnalysisException {
		if (l1.size() != l2.size()) {
			return false;
		}
		for (ExternalRef c1:l1) {
			boolean found = false;
			for (ExternalRef c2:l2) {
				if (c1.equivalent(c2)) {
					found = true;
					break;
				}
			}
			if (!found) {
				return false;
			}
		}
		return true;
	}

	public void testRefTypeToString() {
		Workbook wb = new HSSFWorkbook();
		ExternalRefsSheet.create(wb, "External Refs");
		ExternalRefsSheet externalRefsSheet = new ExternalRefsSheet(wb, "External Refs", modelStore, 
				DOCUMENT_NAMSPACE, copyManager);
		assertEquals(CPE32_NAME, externalRefsSheet.refTypeToString(REFERENCE_TYPE_CPE32));
		assertEquals(MAVEN_NAME, externalRefsSheet.refTypeToString(REFERENCE_TYPE_MAVEN));
		assertEquals(LOCAL_REFERENCE_TYPE_NAME, externalRefsSheet.refTypeToString(REFERENCE_TYPE_LOCAL_TO_PACKAGE));
		assertEquals(FULL_REFRENCE_TYPE_URI, externalRefsSheet.refTypeToString(REFERENCE_TYPE_FULL_URI));
	}
	
	public void testStringToReferenceType() {
		Workbook wb = new HSSFWorkbook();
		ExternalRefsSheet.create(wb, "External Refs");
		ExternalRefsSheet externalRefsSheet = new ExternalRefsSheet(wb, "External Refs", modelStore, 
				DOCUMENT_NAMSPACE, copyManager);
		assertTrue(REFERENCE_TYPE_CPE32.equals(externalRefsSheet.stringToRefType(CPE32_NAME)));
		assertTrue(REFERENCE_TYPE_MAVEN.equals(externalRefsSheet.stringToRefType(MAVEN_NAME)));
		assertTrue(REFERENCE_TYPE_LOCAL_TO_PACKAGE.equals(externalRefsSheet.stringToRefType(LOCAL_REFERENCE_TYPE_NAME)));
		assertTrue(REFERENCE_TYPE_FULL_URI.equals(externalRefsSheet.stringToRefType(FULL_REFRENCE_TYPE_URI)));
	}

}
