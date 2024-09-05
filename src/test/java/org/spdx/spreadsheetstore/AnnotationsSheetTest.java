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

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultModelStore;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.core.ModelRegistry;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.Annotation;
import org.spdx.library.model.v2.SpdxModelInfoV2_X;
import org.spdx.library.model.v2.enumerations.AnnotationType;
import org.spdx.library.model.v3_0_1.SpdxModelInfoV3_0;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

/**
 * @author gary
 *
 */
public class AnnotationsSheetTest extends TestCase {

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
		AnnotationsSheet.create(wb, "Annotations");
		AnnotationsSheet sheet = new AnnotationsSheet(wb, "Annotations", 
				modelStore, DOCUMENT_URI, copyManager);
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}
	
	public void testAdd() throws InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		AnnotationsSheet.create(wb, "Annotations");
		AnnotationsSheet sheet = new AnnotationsSheet(wb, "Annotations", 
				modelStore, DOCUMENT_URI, copyManager);
		
		Annotation an1 = new Annotation(modelStore, DOCUMENT_URI, 
				modelStore.getNextId(IdType.Anonymous), copyManager, true);
		an1.setAnnotator("Person: Annotator1");
		an1.setAnnotationDate("2010-01-29T18:30:22Z");
		an1.setAnnotationType(AnnotationType.OTHER);
		an1.setComment("Comment1");
		
		Annotation an2 = new Annotation(modelStore, DOCUMENT_URI, 
				modelStore.getNextId(IdType.Anonymous), copyManager, true);
		an2.setAnnotator("Person: Annotator2");
		an2.setAnnotationDate("2015-01-29T18:30:22Z");
		an2.setAnnotationType(AnnotationType.REVIEW);
		an2.setComment("Comment2");
		String id1 = "SPDXRef-1";
		String id2 = "SPDXRef-2";
		sheet.add(an1, id1);
		sheet.add(an2, id2);
		Annotation result = sheet.getAnnotation(1);
		assertTrue(an1.equivalent(result));
		assertEquals(id1, sheet.getElmementId(1));
		result = sheet.getAnnotation(2);
		assertTrue(an2.equivalent(result));
		assertEquals(id2, sheet.getElmementId(2));
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}

}
