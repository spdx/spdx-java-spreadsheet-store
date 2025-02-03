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

import java.io.IOException;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultModelStore;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.core.ModelRegistry;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.Checksum;
import org.spdx.library.model.v2.Relationship;
import org.spdx.library.model.v2.SpdxFile;
import org.spdx.library.model.v2.SpdxModelInfoV2_X;
import org.spdx.library.model.v2.SpdxFile.SpdxFileBuilder;
import org.spdx.library.model.v2.SpdxPackage;
import org.spdx.library.model.v2.enumerations.ChecksumAlgorithm;
import org.spdx.library.model.v2.enumerations.FileType;
import org.spdx.library.model.v2.enumerations.RelationshipType;
import org.spdx.library.model.v2.license.AnyLicenseInfo;
import org.spdx.library.model.v2.license.ConjunctiveLicenseSet;
import org.spdx.library.model.v2.license.DisjunctiveLicenseSet;
import org.spdx.library.model.v2.license.ExtractedLicenseInfo;
import org.spdx.library.model.v2.license.SpdxListedLicense;
import org.spdx.library.model.v3_0_1.SpdxModelInfoV3_0;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

/**
 * @author Gary O'Neall
 */
public class PerFileSheetTest extends TestCase {

	static final String[] NONSTD_IDS = new String[] {"LicenseRef-id1", "LicenseRef-id2", "LicenseRef-id3", "LicenseRef-id4"};
	static final String[] NONSTD_TEXTS = new String[] {"text1", "text2", "text3", "text4"};
	static final String[] STD_IDS = new String[] {"AFL-3.0", "CECILL-B", "EUPL-1.0"};
	static final String[] STD_TEXTS = new String[] {"std text1", "std text2", "std text3"};
	static final String[] ATTRIBUTIONA = new String[] {"Att1"};
	static final String[] ATTRIBUTIONB = new String[] {"Att2", "Att3"};

	ExtractedLicenseInfo[] NON_STD_LICENSES;
	SpdxListedLicense[] STANDARD_LICENSES;
	DisjunctiveLicenseSet[] DISJUNCTIVE_LICENSES;
	ConjunctiveLicenseSet[] CONJUNCTIVE_LICENSES;
	
	ConjunctiveLicenseSet COMPLEX_LICENSE;
	
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
		NON_STD_LICENSES = new ExtractedLicenseInfo[NONSTD_IDS.length];
		for (int i = 0; i < NONSTD_IDS.length; i++) {
			NON_STD_LICENSES[i] = new ExtractedLicenseInfo(NONSTD_IDS[i], NONSTD_TEXTS[i]);
		}
		
		STANDARD_LICENSES = new SpdxListedLicense[STD_IDS.length];
		for (int i = 0; i < STD_IDS.length; i++) {
			STANDARD_LICENSES[i] = new SpdxListedLicense("Name "+String.valueOf(i), 
					STD_IDS[i], STD_TEXTS[i], 
					Arrays.asList(new String[] {"URL "+String.valueOf(i), "URL2 "+String.valueOf(i)}), 
					"Notes "+String.valueOf(i), 
					"LicHeader "+String.valueOf(i), "Template "+String.valueOf(i), true, null, null, false, null);
		}
		
		DISJUNCTIVE_LICENSES = new DisjunctiveLicenseSet[3];
		CONJUNCTIVE_LICENSES = new ConjunctiveLicenseSet[2];
		
		DISJUNCTIVE_LICENSES[0] = createDisjunctiveLicenseSet(Arrays.asList(new AnyLicenseInfo[] {
				NON_STD_LICENSES[0], NON_STD_LICENSES[1], STANDARD_LICENSES[1]
		}));
		CONJUNCTIVE_LICENSES[0] = createConjunctiveLicenseSet(Arrays.asList(new AnyLicenseInfo[] {
				STANDARD_LICENSES[0], NON_STD_LICENSES[0], STANDARD_LICENSES[1]
		}));
		CONJUNCTIVE_LICENSES[1] = createConjunctiveLicenseSet(Arrays.asList(new AnyLicenseInfo[] {
				DISJUNCTIVE_LICENSES[0], NON_STD_LICENSES[2]
		}));
		DISJUNCTIVE_LICENSES[1] = createDisjunctiveLicenseSet(Arrays.asList(new AnyLicenseInfo[] {
				CONJUNCTIVE_LICENSES[1], NON_STD_LICENSES[0], STANDARD_LICENSES[0]
		}));
		DISJUNCTIVE_LICENSES[2] = createDisjunctiveLicenseSet(Arrays.asList(new AnyLicenseInfo[] {
				DISJUNCTIVE_LICENSES[1], CONJUNCTIVE_LICENSES[0], STANDARD_LICENSES[2]
		}));
		COMPLEX_LICENSE = createConjunctiveLicenseSet(Arrays.asList(new AnyLicenseInfo[] {
				DISJUNCTIVE_LICENSES[2], NON_STD_LICENSES[2], CONJUNCTIVE_LICENSES[1]
		}));
	}
	
	private DisjunctiveLicenseSet createDisjunctiveLicenseSet(Collection<AnyLicenseInfo> disjunctiveLicenses) throws InvalidSPDXAnalysisException {
		DisjunctiveLicenseSet retval = new DisjunctiveLicenseSet(modelStore, DOCUMENT_URI, modelStore.getNextId(IdType.Anonymous), copyManager, true);
		retval.getMembers().addAll(disjunctiveLicenses);
		return retval;
	}
	
	private ConjunctiveLicenseSet createConjunctiveLicenseSet(Collection<AnyLicenseInfo> conjunctiveLicenses) throws InvalidSPDXAnalysisException {
		ConjunctiveLicenseSet retval = new ConjunctiveLicenseSet(modelStore, DOCUMENT_URI, modelStore.getNextId(IdType.Anonymous), copyManager, true);
		retval.getMembers().addAll(conjunctiveLicenses);
		return retval;
	}

	/* (non-Javadoc)
	 * @see junit.framework.TestCase#tearDown()
	 */
	protected void tearDown() throws Exception {
		super.tearDown();
	}
	
	public void testCreate() throws IOException {

		Workbook wb = new HSSFWorkbook();
		PerFileSheet.create(wb, "File Info");
		PerFileSheet fileInfo = PerFileSheet.openVersion(wb, "File Info", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		String ver = fileInfo.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}
	
	public void testAddAndGet() throws SpreadsheetException, InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		PerFileSheet.create(wb, "File Info");
		PerFileSheet fileInfoSheet = PerFileSheet.openVersion(wb, "File Info", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		AnyLicenseInfo[] testLicenses1 = new AnyLicenseInfo[] {COMPLEX_LICENSE};
		AnyLicenseInfo[] testLicenses2 = new AnyLicenseInfo[] {NON_STD_LICENSES[0]};

		String fileComment1 = "comment 1";
		String[] contributors1 = new String[] {"Contrib1", "Contrib2"};
		String noticeText1 = "notice 1";
		SpdxFile testFile1 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File1", copyManager, 
				"FileName1", COMPLEX_LICENSE, Arrays.asList(testLicenses2), "copyright (c) 1", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment(fileComment1)
				.setLicenseComments("license comments 1")
				.addFileType(FileType.BINARY)
				.setFileContributors(Arrays.asList(contributors1))
				.setNoticeText(noticeText1)
				.setAttributionText(Arrays.asList(ATTRIBUTIONA))
				.build();


		SpdxFile testFile2 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File2", copyManager, 
				"FileName2",  NON_STD_LICENSES[0], Arrays.asList(testLicenses1),  "copyright (c) 12", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment(fileComment1)
				.setLicenseComments("license comments2")
				.addFileType(FileType.SOURCE)
				.setFileContributors(Arrays.asList(contributors1))
				.setNoticeText(noticeText1)
				.setAttributionText(Arrays.asList(ATTRIBUTIONB))
				.build();

		SpdxFile testFile3 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File3", copyManager, 
				"FileName3",  NON_STD_LICENSES[0], Arrays.asList(testLicenses1),  "copyright (c) 123", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment("Comment3")
				.setLicenseComments("license comments3")
				.addFileType(FileType.OTHER)
				.addFileContributor("c1")
				.setNoticeText("Notice")
				.setAttributionText(Arrays.asList(ATTRIBUTIONB))
				.build();
		
		fileInfoSheet.add(testFile1, "SPDXRef-Package1");
		fileInfoSheet.add(testFile2, "SPDXRef-Package1");
		fileInfoSheet.add(testFile3, "SPDXRef-Package1");
		SpdxFile result1 = fileInfoSheet.getFileInfo(1);
		SpdxFile result2 = fileInfoSheet.getFileInfo(2);
		SpdxFile result3 = fileInfoSheet.getFileInfo(3);
		SpdxFile result4 = fileInfoSheet.getFileInfo(4);
		assertTrue(testFile1.equivalent(result1));
		assertTrue(testFile2.equivalent(result2));
		assertTrue(testFile3.equivalent(result3));
		if (result4 != null) {
			fail("expected null");
		}
	}
	
	public void testCsv() {
		List<String> strings = Arrays.asList(new String[] {"Test1", "\"Quoted test2\"", "", "Test4 with, comma"});
		String csvString = PerFileSheet.stringsToCsv(strings);
		List<String> result = PerFileSheet.csvToStrings(csvString);
		assertEquals(strings.size(), result.size());
		for (int i = 0; i < strings.size(); i++) {
			assertEquals(strings.get(i), result.get(i));
		}
	}
	
	public void testArtifactOf() throws InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		PerFileSheet.create(wb, "File Info");
		PerFileSheet fileInfoSheet = PerFileSheet.openVersion(wb, "File Info", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		AnyLicenseInfo[] testLicenses2 = new AnyLicenseInfo[] {NON_STD_LICENSES[0]};

		String fileComment1 = "comment 1";
		String[] contributors1 = new String[] {"Contrib1", "Contrib2"};
		String noticeText1 = "notice 1";
		SpdxFile testFile1 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File1", copyManager, 
				"FileName1", COMPLEX_LICENSE, Arrays.asList(testLicenses2), "copyright (c) 1", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment(fileComment1)
				.setLicenseComments("license comments 1")
				.addFileType(FileType.BINARY)
				.setFileContributors(Arrays.asList(contributors1))
				.setNoticeText(noticeText1)
				.setAttributionText(Arrays.asList(ATTRIBUTIONA))
				.build();
		
		fileInfoSheet.add(testFile1, "SPDXRef-Package1");
		Row row = fileInfoSheet.sheet.getRow(1);
		List<String> projectNames = Arrays.asList(new String[] {"projectA", "projectB"});
		List<String> projectHomePages = Arrays.asList(new String[] {"http://www.projectA", "http://www.projectB"});
		row.createCell(PerFileSheetV2d2.ARTIFACT_OF_PROJECT_COL).setCellValue(PerFileSheetV2d2.stringsToCsv(projectNames));
		row.createCell(PerFileSheetV2d2.ARTIFACT_OF_HOMEPAGE_COL).setCellValue(PerFileSheetV2d2.stringsToCsv(projectHomePages));
		SpdxFile result = fileInfoSheet.getFileInfo(1);
		Collection<Relationship> relationships = result.getRelationships();
		assertEquals(2, relationships.size());
		for (Relationship rel:relationships) {
			assertTrue(rel.getRelatedSpdxElement().get() instanceof SpdxPackage);
			assertEquals(RelationshipType.GENERATED_FROM, rel.getRelationshipType());
			SpdxPackage pkg = (SpdxPackage)rel.getRelatedSpdxElement().get();
			if (pkg.getName().get().equals("projectA")) {
				assertEquals("http://www.projectA", pkg.getHomepage().get());
			} else if (pkg.getName().get().equals("projectB")) {
				assertEquals("http://www.projectB", pkg.getHomepage().get());
			}
		}
	}
	
	public void testFileDependency() throws InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		PerFileSheet.create(wb, "File Info");
		PerFileSheet fileInfoSheet = PerFileSheet.openVersion(wb, "File Info", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		AnyLicenseInfo[] testLicenses1 = new AnyLicenseInfo[] {COMPLEX_LICENSE};
		AnyLicenseInfo[] testLicenses2 = new AnyLicenseInfo[] {NON_STD_LICENSES[0]};

		String fileComment1 = "comment 1";
		String[] contributors1 = new String[] {"Contrib1", "Contrib2"};
		String noticeText1 = "notice 1";
		SpdxFile testFile1 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File1", copyManager, 
				"FileName1", COMPLEX_LICENSE, Arrays.asList(testLicenses2), "copyright (c) 1", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment(fileComment1)
				.setLicenseComments("license comments 1")
				.addFileType(FileType.BINARY)
				.setFileContributors(Arrays.asList(contributors1))
				.setNoticeText(noticeText1)
				.setAttributionText(Arrays.asList(ATTRIBUTIONA))
				.build();


		SpdxFile testFile2 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File2", copyManager, 
				"FileName2",  NON_STD_LICENSES[0], Arrays.asList(testLicenses1),  "copyright (c) 12", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment(fileComment1)
				.setLicenseComments("license comments2")
				.addFileType(FileType.SOURCE)
				.setFileContributors(Arrays.asList(contributors1))
				.setNoticeText(noticeText1)
				.setAttributionText(Arrays.asList(ATTRIBUTIONB))
				.build();

		SpdxFile testFile3 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File3", copyManager, 
				"FileName3",  NON_STD_LICENSES[0], Arrays.asList(testLicenses1),  "copyright (c) 123", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment("Comment3")
				.setLicenseComments("license comments3")
				.addFileType(FileType.OTHER)
				.addFileContributor("c1")
				.setNoticeText("Notice")
				.setAttributionText(Arrays.asList(ATTRIBUTIONB))
				.build();
		
		fileInfoSheet.add(testFile1, "SPDXRef-Package1");
		fileInfoSheet.add(testFile2, "SPDXRef-Package1");
		fileInfoSheet.add(testFile3, "SPDXRef-Package1");
		
		Row row = fileInfoSheet.sheet.getRow(1);
		row.createCell(PerFileSheetV2d2.FILE_DEPENDENCIES_COL).setCellValue("FileName2, FileName3");
		
		SpdxFile result = fileInfoSheet.getFileInfo(1);
		Collection<Relationship> relationships = result.getRelationships();
		assertEquals(2, relationships.size());
		for (Relationship rel:relationships) {
			assertTrue(rel.getRelatedSpdxElement().get() instanceof SpdxFile);
			assertEquals(RelationshipType.DEPENDS_ON, rel.getRelationshipType());
			assertTrue(rel.getRelatedSpdxElement().get().getId().equals("SPDXRef-File2") ||
					rel.getRelatedSpdxElement().get().getId().equals("SPDXRef-File3"));
		}
	}


}
