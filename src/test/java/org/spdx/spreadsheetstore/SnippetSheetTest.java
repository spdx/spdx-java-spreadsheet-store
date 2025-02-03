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

import java.util.Arrays;
import java.util.Collection;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.core.DefaultModelStore;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.core.ModelRegistry;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.model.v2.Checksum;
import org.spdx.library.model.v2.SpdxFile;
import org.spdx.library.model.v2.SpdxModelInfoV2_X;
import org.spdx.library.model.v2.SpdxFile.SpdxFileBuilder;
import org.spdx.library.model.v2.SpdxSnippet;
import org.spdx.library.model.v2.SpdxSnippet.SpdxSnippetBuilder;
import org.spdx.library.model.v2.enumerations.ChecksumAlgorithm;
import org.spdx.library.model.v2.enumerations.FileType;
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
public class SnippetSheetTest extends TestCase {

	static final String[] NONSTD_IDS = new String[] {"LicenseRef-id1", "LicenseRef-id2", "LicenseRef-id3", "LicenseRef-id4"};
	static final String[] NONSTD_TEXTS = new String[] {"text1", "text2", "text3", "text4"};
	static final String[] STD_IDS = new String[] {"AFL-3.0", "CECILL-B", "EUPL-1.0"};
	static final String[] STD_TEXTS = new String[] {"std text1", "std text2", "std text3"};

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
	
	public void testCreate() {
		Workbook wb = new HSSFWorkbook();
		SnippetSheet.create(wb, "Snippets");
		SnippetSheet sheet = new SnippetSheet(wb, "Snippets", 
				modelStore, DOCUMENT_URI, copyManager);
		String ver = sheet.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}

	public void testAddGet() throws SpreadsheetException, InvalidSPDXAnalysisException {
		Workbook wb = new HSSFWorkbook();
		SnippetSheet.create(wb, "Snippets");
		SnippetSheet snippetSheet = new SnippetSheet(wb, "Snippets", 
				modelStore, DOCUMENT_URI, copyManager);
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
				.build();


		SpdxFile testFile2 = new SpdxFileBuilder(modelStore, DOCUMENT_URI, "SPDXRef-File2", copyManager, 
				"FileName2",  NON_STD_LICENSES[0], Arrays.asList(testLicenses1),  "copyright (c) 12", 
				Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.setComment(fileComment1)
				.setLicenseComments("license comments2")
				.addFileType(FileType.SOURCE)
				.setFileContributors(Arrays.asList(contributors1))
				.setNoticeText(noticeText1)
				.build();

		SpdxSnippet snippet1 = new SpdxSnippetBuilder(modelStore, 
				DOCUMENT_URI, "SPDXRef-snippet1", copyManager, "snippet1", 
				COMPLEX_LICENSE, Arrays.asList(testLicenses1), "copyright (c) 1", testFile1, 5, 10)
				.setLineRange(5, 10)
				.setComment("comment1")
				.setLicenseComments("license comments 1")
				.build();

		SpdxSnippet snippet2 = new SpdxSnippetBuilder(modelStore, 
				DOCUMENT_URI, "SPDXRef-snippet2", copyManager, "snippet2", 
				NON_STD_LICENSES[0], Arrays.asList(testLicenses2), "copyright (c) 2", testFile2, 7, 8)
				.setLineRange(55, 1213)
				.setComment("comment2")
				.setLicenseComments("license comments 2")
				.build();
		
		snippetSheet.add(snippet1);
		snippetSheet.add(snippet2);
		
		SpdxSnippet result1 = snippetSheet.getSnippet(1);
		SpdxSnippet result2 = snippetSheet.getSnippet(2);
		SpdxSnippet result3 = snippetSheet.getSnippet(3);
		assertTrue(snippet1.equivalent(result1));
		assertTrue(snippet2.equivalent(result2));
		assertTrue(result3 == null);
		assertEquals(snippet1.getSnippetFromFile().getId(), snippetSheet.getSnippetFileId(1));
		assertEquals(snippet2.getSnippetFromFile().getId(), snippetSheet.getSnippetFileId(2));
		assertTrue(snippetSheet.verify() == null);
	}
}
