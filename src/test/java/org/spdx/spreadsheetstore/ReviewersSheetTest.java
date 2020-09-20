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

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Collection;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxConstants;
import org.spdx.library.model.Annotation;
import org.spdx.library.model.SpdxDocument;
import org.spdx.library.model.SpdxModelFactory;
import org.spdx.library.model.enumerations.AnnotationType;
import org.spdx.storage.IModelStore;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

/**
 * @author gary
 *
 */
public class ReviewersSheetTest extends TestCase {

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
		SpdxModelFactory.createSpdxDocument(modelStore, DOCUMENT_URI, copyManager);
	}

	/* (non-Javadoc)
	 * @see junit.framework.TestCase#tearDown()
	 */
	protected void tearDown() throws Exception {
		super.tearDown();
	}
	
	public void testCreate() {
		Workbook wb = new HSSFWorkbook();
		ReviewersSheet.create(wb, "Review");
		ReviewersSheet sheet = new ReviewersSheet(wb, "Review", 
				modelStore, DOCUMENT_URI, copyManager);
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}
	
	public void testAddReviewsToDocAnnotations() throws InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		ReviewersSheet.create(wb, "Review");
		ReviewersSheet sheet = new ReviewersSheet(wb, "Review", 
				modelStore, DOCUMENT_URI, copyManager);
		Calendar cal = Calendar.getInstance();
		cal.set(2020, 5, 18);
		Date date1 = new Date();
		Date date2 = cal.getTime();
		String annotator1 = "Person: Annotator1";
		String annotator2 = "Person: Annotator2";
		String comment1 = "comment1";
		String comment2 = "comment2";
		
		Row row1 = sheet.addRow();
		row1.createCell(ReviewersSheet.REVIEWER_COL).setCellValue(annotator1);
		Cell dateCell = row1.createCell(ReviewersSheet.TIMESTAMP_COL);
		dateCell.setCellValue(date1);
		dateCell.setCellStyle(sheet.dateStyle);
		row1.createCell(ReviewersSheet.COMMENT_COL).setCellValue(comment1);
		
		Row row2 = sheet.addRow();
		row2.createCell(ReviewersSheet.REVIEWER_COL).setCellValue(annotator2);
		Cell dateCell2 = row2.createCell(ReviewersSheet.TIMESTAMP_COL);
		dateCell2.setCellValue(date2);
		dateCell2.setCellStyle(sheet.dateStyle);
		row2.createCell(ReviewersSheet.COMMENT_COL).setCellValue(comment2);
		
		sheet.addReviewsToDocAnnotations();
		
		SimpleDateFormat format = new SimpleDateFormat(SpdxConstants.SPDX_DATE_FORMAT);
		Collection<Annotation> docAnnotations = new SpdxDocument(modelStore, DOCUMENT_URI, copyManager, false).getAnnotations();
		assertEquals(2, docAnnotations.size());
		docAnnotations.forEach(annotation -> {
			try {
				if (annotation.getAnnotator().equals(annotator1)) {
					assertEquals(comment1, annotation.getComment());
					assertEquals(AnnotationType.REVIEW, annotation.getAnnotationType());
					assertEquals(format.format(date1), annotation.getAnnotationDate());
				} else {
					assertEquals(annotator2, annotation.getAnnotator());
					assertEquals(comment2, annotation.getComment());
					assertEquals(AnnotationType.REVIEW, annotation.getAnnotationType());
					assertEquals(format.format(date2), annotation.getAnnotationDate());
				}
			} catch (InvalidSPDXAnalysisException ex) {
				fail(ex.getMessage());
			}
		});
	}

}
