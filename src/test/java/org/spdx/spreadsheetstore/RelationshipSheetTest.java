/*
 * SPDX-FileContributor: Gary O'Neall
 * SPDX-FileCopyrightText: Copyright (c) 2020 Source Auditor Inc.
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 * <p>
 *   Licensed under the Apache License, Version 2.0 (the "License");
 *   you may not use this file except in compliance with the License.
 *   You may obtain a copy of the License at
 * <p>
 *       https://www.apache.org/licenses/LICENSE-2.0
 * <p>
 *   Unless required by applicable law or agreed to in writing, software
 *   distributed under the License is distributed on an "AS IS" BASIS,
 *   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 *   See the License for the specific language governing permissions and
 *   limitations under the License.
 */
package org.spdx.spreadsheetstore;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultModelStore;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.core.ModelRegistry;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.GenericSpdxElement;
import org.spdx.library.model.v2.Relationship;
import org.spdx.library.model.v2.SpdxElement;
import org.spdx.library.model.v2.SpdxModelInfoV2_X;
import org.spdx.library.model.v2.enumerations.RelationshipType;
import org.spdx.library.model.v3_0_1.SpdxModelInfoV3_0;
import org.spdx.storage.IModelStore;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

/**
 * @author Gary O'Neall
 */
public class RelationshipSheetTest extends TestCase {

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
		RelationshipsSheet.create(wb, "Relationship Info");
		RelationshipsSheet sheet = new RelationshipsSheet(wb, "Relationship Info", 
				modelStore, DOCUMENT_URI, copyManager);
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}
	
	public void testAddandGet() throws InvalidSPDXAnalysisException, SpreadsheetException {
		SpdxElement element1 = new GenericSpdxElement(modelStore, DOCUMENT_URI, "SPDXRef-1", copyManager, true);
		SpdxElement element2 = new GenericSpdxElement(modelStore, DOCUMENT_URI, "SPDXRef-2", copyManager, true);		
		Relationship rel1 = element2.createRelationship(element1, RelationshipType.AMENDS, "Comment1");
		SpdxElement element3 = new GenericSpdxElement(modelStore, DOCUMENT_URI, "SPDXRef-3", copyManager, true);		
		Relationship rel2 = element3.createRelationship(element2, RelationshipType.CONTAINS, null);
		Relationship rel3 = element1.createRelationship(element3, RelationshipType.SPECIFICATION_FOR, "Comment2");
		Workbook wb = new HSSFWorkbook();
		RelationshipsSheet.create(wb, "Relationship Info");
		RelationshipsSheet sheet = new RelationshipsSheet(wb, "Relationship Info", 
				modelStore, DOCUMENT_URI, copyManager);
		String id1 = "SPDXRef-first";
		String id2 = "SPDXRef-second";
		String id3 = "SPDXRef-third";
		sheet.add(rel1, id1);
		sheet.add(rel2, id2);
		sheet.add(rel3, id3);
		Relationship result = sheet.getRelationship(1);
		assertTrue(result.equivalent(rel1));
		assertEquals(id1, sheet.getElmementId(1));
		result = sheet.getRelationship(2);
		assertTrue(result.equivalent(rel2));
		assertEquals(id2, sheet.getElmementId(2));
		result = sheet.getRelationship(3);
		assertTrue(result.equivalent(rel3));
		assertEquals(id3, sheet.getElmementId(3));
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}

}
