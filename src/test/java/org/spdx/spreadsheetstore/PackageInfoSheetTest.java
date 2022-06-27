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

import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Collection;
import java.util.GregorianCalendar;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxConstants;
import org.spdx.library.model.Checksum;
import org.spdx.library.model.SpdxPackage;
import org.spdx.library.model.SpdxPackageVerificationCode;
import org.spdx.library.model.SpdxPackage.SpdxPackageBuilder;
import org.spdx.library.model.enumerations.ChecksumAlgorithm;
import org.spdx.library.model.enumerations.Purpose;
import org.spdx.library.model.license.AnyLicenseInfo;
import org.spdx.library.model.license.ConjunctiveLicenseSet;
import org.spdx.library.model.license.DisjunctiveLicenseSet;
import org.spdx.library.model.license.ExtractedLicenseInfo;
import org.spdx.library.model.license.SpdxNoneLicense;
import org.spdx.storage.IModelStore;
import org.spdx.storage.IModelStore.IdType;
import org.spdx.storage.simple.InMemSpdxStore;

import junit.framework.TestCase;

/**
 * @author gary
 *
 */
public class PackageInfoSheetTest extends TestCase {

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
	
	public void testCreate() throws IOException {

		Workbook wb = new HSSFWorkbook();
		PackageInfoSheet.create(wb, "Package Info");
		PackageInfoSheet pkgInfo = PackageInfoSheet.openVersion(wb, "Package Info", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		String ver = pkgInfo.verify();
		if (ver != null && !ver.isEmpty()){
			fail(ver);
		}
	}


	public void testAddAndGet() throws SpreadsheetException, InvalidSPDXAnalysisException {
		Collection<AnyLicenseInfo> testLicenses1 = new ArrayList<>();
		testLicenses1.add(createExtractedLicense("License1", "License1Text"));
		Collection<AnyLicenseInfo> disjunctiveLicenses = new ArrayList<>();
		disjunctiveLicenses.add(createExtractedLicense("disj1", "disj1 Text"));
		disjunctiveLicenses.add(createExtractedLicense("disj2", "disj2 Text"));
		disjunctiveLicenses.add(createExtractedLicense("disj3", "disj3 Text"));
		testLicenses1.add(createDisjunctiveLicenseSet(disjunctiveLicenses));
		Collection<AnyLicenseInfo> conjunctiveLicenses = new ArrayList<>();
		conjunctiveLicenses.add(createExtractedLicense("conj1", "conj1 Text"));
		conjunctiveLicenses.add(createExtractedLicense("conj2", "conj2 Text"));

		testLicenses1.add(createConjunctiveLicenseSet(conjunctiveLicenses));
		AnyLicenseInfo testLicense1 = createDisjunctiveLicenseSet(testLicenses1);

//		String lic1String = PackageInfoSheet.licensesToString(testLicenses1);
		Collection<AnyLicenseInfo>  testLicenses2 =  new ArrayList<>();
		testLicenses2.add(createExtractedLicense("License3", "License 3 text"));
		testLicenses2.add(createExtractedLicense("License4", "License 4 text"));
		AnyLicenseInfo testLicense2 = createConjunctiveLicenseSet(testLicenses2);
		Collection<AnyLicenseInfo> testLicenseInfos = new ArrayList<>();
		testLicenseInfos.add(new SpdxNoneLicense());
		SpdxPackageVerificationCode testVerification = new SpdxPackageVerificationCode(modelStore, DOCUMENT_URI, modelStore.getNextId(IdType.Anonymous, DOCUMENT_URI), copyManager, true);
		testVerification.setValue("value");
		testVerification.getExcludedFileNames().add("skippedfil1");
		testVerification.getExcludedFileNames().add("skippedfile2");
//		String lic2String = PackageInfoSheet.licensesToString(testLicenses2);

		Checksum sha1 = Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12");
		Checksum blake2b = Checksum.create(modelStore, DOCUMENT_URI, ChecksumAlgorithm.BLAKE2b_384, "aaabd89c926ab525c242e6621f2f5fa73aa4afe3d9e24aed727faaadd6af38b620bdb623dd2b4788b1c8086984af8706");

		String releaseDate = new SimpleDateFormat(SpdxConstants.SPDX_DATE_FORMAT).format(new GregorianCalendar(2021, Calendar.JANUARY, 11).getTime());
		String buildDate = new SimpleDateFormat(SpdxConstants.SPDX_DATE_FORMAT).format(new GregorianCalendar(2020, Calendar.JANUARY, 11).getTime());
		String validUntilDate = new SimpleDateFormat(SpdxConstants.SPDX_DATE_FORMAT).format(new GregorianCalendar(2023, Calendar.JANUARY, 11).getTime());
		SpdxPackage pkgInfo1 = new SpdxPackageBuilder(modelStore, DOCUMENT_URI,  "SPDXRef-Package1", copyManager, 
				"decname1", testLicense1, "dec-copyright1", testLicense2)
				.addChecksum(sha1)
				.setComment("Comment1")
				.setLicenseInfosFromFile(testLicenseInfos)
				.setLicenseComments("license comments")
				.setDescription("desc1")
				.setDownloadLocation("http://www.spdx.org/download")
				.setHomepage("http://www.home.page1")
				.setOriginator("Organization: originator1")
				.setPackageFileName("machinename1")
				.setPackageVerificationCode(testVerification)
				.setSourceInfo("sourceinfo1")
				.setSummary("short desc1")
				.setSupplier("Person: supplier1")
				.setVersionInfo("Version1")
				.setFilesAnalyzed(true)
				.setAttributionText(Arrays.asList(new String[]{"Att1", "att2"}))
				.setPrimaryPurpose(Purpose.CONTAINER)
				.setReleaseDate(releaseDate)
				.setBuiltDate(buildDate)
				.setValidUntilDate(validUntilDate)
				.addChecksum(blake2b)
				.build();

		SpdxPackage pkgInfo2 =  new SpdxPackageBuilder(modelStore, DOCUMENT_URI,  "SPDXRef-Package2", copyManager, 
				"decname1", testLicense1, "dec-copyright1", testLicense2)
				.addChecksum(sha1)
				.setLicenseInfosFromFile(testLicenseInfos)
				.setComment("Comment1")
				.setLicenseComments("license comments2")
				.setDescription("desc1")
				.setDownloadLocation("http://www.spdx.org/download")
				.setHomepage("http://www.home.page2")
				.setOriginator("Organization: originator1")
				.setPackageFileName("machinename1")
				.setPackageVerificationCode(testVerification)
				.setSourceInfo("sourceinfo1")
				.setSummary("short desc1")
				.setSupplier("Person: supplier1")
				.setVersionInfo("Version2")
				.setFilesAnalyzed(false)
				.addAttributionText("Att3")
				.build();

		Workbook wb = new HSSFWorkbook();
		PackageInfoSheet.create(wb, "Package Info");
		PackageInfoSheet pkgInfoSheet = PackageInfoSheet.openVersion(wb, "Package Info", 
				SpdxSpreadsheet.CURRENT_VERSION, modelStore, DOCUMENT_URI, copyManager);
		pkgInfoSheet.add(pkgInfo1);
		pkgInfoSheet.add(pkgInfo2);
		SpdxPackage tstPkgInfo1 = pkgInfoSheet.getPackages().get(0);
		SpdxPackage tstPkgInfo2 = pkgInfoSheet.getPackages().get(1);
		assertTrue(pkgInfo1.equivalent(tstPkgInfo1));
		assertEquals(pkgInfo1.getId(), tstPkgInfo1.getId());
		assertEquals(Purpose.CONTAINER, tstPkgInfo1.getPrimaryPurpose().get());
		assertEquals(releaseDate, tstPkgInfo1.getReleaseDate().get());
		assertEquals(buildDate, tstPkgInfo1.getBuiltDate().get());
		assertEquals(validUntilDate, tstPkgInfo1.getValidUntilDate().get());
		assertTrue(pkgInfo2.equivalent(tstPkgInfo2));
		assertEquals(pkgInfo2.getId(), tstPkgInfo2.getId());
		assertEquals(2, pkgInfoSheet.getPackages().size());
	}

	private DisjunctiveLicenseSet createDisjunctiveLicenseSet(Collection<AnyLicenseInfo> disjunctiveLicenses) throws InvalidSPDXAnalysisException {
		DisjunctiveLicenseSet retval = new DisjunctiveLicenseSet(modelStore, DOCUMENT_URI, modelStore.getNextId(IdType.Anonymous, DOCUMENT_URI), copyManager, true);
		retval.getMembers().addAll(disjunctiveLicenses);
		return retval;
	}
	
	private ConjunctiveLicenseSet createConjunctiveLicenseSet(Collection<AnyLicenseInfo> conjunctiveLicenses) throws InvalidSPDXAnalysisException {
		ConjunctiveLicenseSet retval = new ConjunctiveLicenseSet(modelStore, DOCUMENT_URI, modelStore.getNextId(IdType.Anonymous, DOCUMENT_URI), copyManager, true);
		retval.getMembers().addAll(conjunctiveLicenses);
		return retval;
	}

	private ExtractedLicenseInfo createExtractedLicense(String id, String text) throws InvalidSPDXAnalysisException {
		ExtractedLicenseInfo retval = new ExtractedLicenseInfo(modelStore, DOCUMENT_URI, id, copyManager, true);
		retval.setExtractedText(text);
		return retval;
	}

}
