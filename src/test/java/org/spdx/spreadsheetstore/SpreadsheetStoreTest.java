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

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Stream;

import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxConstants;
import org.spdx.library.model.Annotation;
import org.spdx.library.model.Checksum;
import org.spdx.library.model.ExternalDocumentRef;
import org.spdx.library.model.ExternalSpdxElement;
import org.spdx.library.model.ReferenceType;
import org.spdx.library.model.Relationship;
import org.spdx.library.model.SpdxCreatorInformation;
import org.spdx.library.model.SpdxDocument;
import org.spdx.library.model.SpdxElement;
import org.spdx.library.model.SpdxFile;
import org.spdx.library.model.SpdxModelFactory;
import org.spdx.library.model.SpdxNoAssertionElement;
import org.spdx.library.model.SpdxPackage;
import org.spdx.library.model.SpdxSnippet;
import org.spdx.library.model.enumerations.AnnotationType;
import org.spdx.library.model.enumerations.ChecksumAlgorithm;
import org.spdx.library.model.enumerations.FileType;
import org.spdx.library.model.enumerations.ReferenceCategory;
import org.spdx.library.model.enumerations.RelationshipType;
import org.spdx.library.model.license.AnyLicenseInfo;
import org.spdx.library.model.license.ExtractedLicenseInfo;
import org.spdx.library.model.license.LicenseInfoFactory;
import org.spdx.library.model.license.ListedLicenses;
import org.spdx.library.model.license.SpdxNoAssertionLicense;
import org.spdx.library.referencetype.ListedReferenceTypes;
import org.spdx.storage.simple.InMemSpdxStore;
import org.spdx.utility.compare.SpdxCompareException;
import org.spdx.utility.compare.SpdxComparer;

import junit.framework.TestCase;

/**
 * @author gary
 *
 */
public class SpreadsheetStoreTest extends TestCase {
	
	private static final String SPREADSHEET_2_0_FILENAME = "TestFiles" + File.separator + "SPDXSpreadsheetExample-2.0.xlsx";
	private static final String SPREADSHEET_2_2_FILENAME = "TestFiles" + File.separator + "SPDXSpreadsheetExample-v2.2.xlsx";
	private static final String SPREADSHEET_2_2_FILENAME_XLS = "TestFiles" + File.separator + "SPDXSpreadsheetExample-v2.2.xls";

	private static final String LICENSEREF1_TEXT = "/*\n"+
			" * (c) Copyright 2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009 Hewlett-Packard Development Company, LP\n"+
			" * All rights reserved.\n"+
			" *\n"+
			" * Redistribution and use in source and binary forms, with or without\n"+
			" * modification, are permitted provided that the following conditions\n"+
			" * are met:\n"+
			" * 1. Redistributions of source code must retain the above copyright\n"+
			" *    notice, this list of conditions and the following disclaimer.\n"+
			" * 2. Redistributions in binary form must reproduce the above copyright\n"+
			" *    notice, this list of conditions and the following disclaimer in the\n"+
			" *    documentation and/or other materials provided with the distribution.\n"+
			" * 3. The name of the author may not be used to endorse or promote products\n"+
			" *    derived from this software without specific prior written permission.\n"+
			" *\n"+
			" * THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR\n"+
			" * IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES\n"+
			" * OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.\n"+
			" * IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT,\n"+
			" * INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT\n"+
			" * NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,\n"+
			" * DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY\n"+
			" * THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT\n"+
			" * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF\n"+
			" * THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\n"+
			"*/";
	private static final String LICENSEREF2_TEXT = "This package includes the GRDDL parser developed by Hewlett Packard under the following license:\n"+
			"� Copyright 2007 Hewlett-Packard Development Company, LP\n"+
			"\n"+
			"Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met: \n"+
			"\n"+
			"Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer. \n"+
			"Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution. \n"+
			"The name of the author may not be used to endorse or promote products derived from this software without specific prior written permission. \n"+
			"THIS SOFTWARE IS PROVIDED BY THE AUTHOR ``AS IS'' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE AUTHOR BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.";

	private static final String LICENSEREF3_TEXT = "The CyberNeko Software License, Version 1.0\n"+
					"\n"+
					"\n"+
		"(C) Copyright 2002-2005, Andy Clark.  All rights reserved.\n"+
		"\n"+
		"Redistribution and use in source and binary forms, with or without\n"+
		"modification, are permitted provided that the following conditions\n"+
		"are met:\n"+
		"\n"+
		"1. Redistributions of source code must retain the above copyright\n"+
		"   notice, this list of conditions and the following disclaimer. \n"+
		   "\n"+
		"2. Redistributions in binary form must reproduce the above copyright\n"+
		"   notice, this list of conditions and the following disclaimer in\n"+
		"   the documentation and/or other materials provided with the\n"+
		"   distribution.\n"+
		   "\n"+
		"3. The end-user documentation included with the redistribution,\n"+
		"   if any, must include the following acknowledgment:  \n"+
		"     \"This product includes software developed by Andy Clark.\"\n"+
		"   Alternately, this acknowledgment may appear in the software itself,\n"+
		"   if and wherever such third-party acknowledgments normally appear.\n"+
		   "\n"+
		"4. The names \"CyberNeko\" and \"NekoHTML\" must not be used to endorse\n"+
		"   or promote products derived from this software without prior \n"+
		"   written permission. For written permission, please contact \n"+
		"   andyc@cyberneko.net.\n"+
		   "\n"+
		"5. Products derived from this software may not be called \"CyberNeko\",\n"+
		"   nor may \"CyberNeko\" appear in their name, without prior written\n"+
		"   permission of the author.\n"+
		   "\n"+
		"THIS SOFTWARE IS PROVIDED ``AS IS'' AND ANY EXPRESSED OR IMPLIED\n"+
		"WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES\n"+
		"OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE\n"+
		"DISCLAIMED.  IN NO EVENT SHALL THE AUTHOR OR OTHER CONTRIBUTORS\n"+
		"BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, \n"+
		"OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT \n"+
		"OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR \n"+
		"BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, \n"+
		"WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE \n"+
		"OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, \n"+
		"EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.";
	private static final String LICENSEREF4_TEXT = "/*\n"+
			" * (c) Copyright 2009 University of Bristol\n"+
			" * All rights reserved.\n"+
			" *";
	private static final String LICENSEREF_BEER_TEXT = "\"THE BEER-WARE LICENSE\" (Revision 42):\n"+
			"phk@FreeBSD.ORG wrote this file. As long as you retain this notice you\n"+
			"can do whatever you want with this stuff. If we meet some day, and you think this stuff is worth it, you can buy me a beer in return Poul-Henning Kamp  </\n"+
			"LicenseName: Beer-Ware License (Version 42)\n"+
			"LicenseCrossReference:  http://people.freebsd.org/~phk/\n"+
			"LicenseComment: \n"+
			"The beerware license has a couple of other standard variants.";
	private static final String SPDXREF_FILE_NOTICE = "Copyright (c) 2001 Aaron Lehmann aaroni@vitelus.com\n"+
			"\n"+
			"Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the �Software�), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions: \n"+
			"The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\n"+
			"\n"+
			"THE SOFTWARE IS PROVIDED �AS IS', WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.  IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
	private static final String SPDXREF_COMMONS_NOTICE = "Apache Commons Lang\n"+
			"Copyright 2001-2011 The Apache Software Foundation\n"+
			"\n"+
			"This product includes software developed by\n"+
			"The Apache Software Foundation (http://www.apache.org/).\n"+
			"\n"+
			"This product includes software from the Spring Framework,\n"+
			"under the Apache License 2.0 (see: StringUtils.containsWhitespace())";		
	
	private Map<String, SpdxPackage> comparePackages = new HashMap<>();
	private Map<String, SpdxFile> compareFiles = new HashMap<>();
	private Map<String, SpdxSnippet> compareSnippets = new HashMap<>();
	private SpdxDocument compareDocument;
	private InMemSpdxStore compareStore;
	
	/* (non-Javadoc)
	 * @see junit.framework.TestCase#setUp()
	 */
	protected void setUp() throws Exception {
		super.setUp();
		compareStore = new InMemSpdxStore();
		String compDocUri = "http://compare.doc.uri/test";
		ModelCopyManager copyManager = new ModelCopyManager();
		compareDocument = SpdxModelFactory.createSpdxDocument(compareStore, compDocUri, copyManager);
		compareDocument.setName("SPDX-Tools-v2.0");
		compareDocument.setComment("This document was created using SPDX 2.0 using licenses from the web site.");
		
		compareDocument.setCreationInfo(compareDocument.createCreationInfo(Arrays.asList(new String[] {"Tool: LicenseFind-1.0", "Organization: ExampleCodeInspect ()", "Person: Jane Doe ()"}), 
				"2010-01-29T18:30:22Z"));
		compareDocument.getCreationInfo().setLicenseListVersion("3.8");
		compareDocument.getCreationInfo().setComment("This package has been shipped in source and binary form.\n"+
														"The binaries were created with gcc 4.5.1 and expect to link to\n"+
														"compatible system run time libraries.");
		AnyLicenseInfo noAssertionLicense = new SpdxNoAssertionLicense();
		AnyLicenseInfo gpl2only = ListedLicenses.getListedLicenses().getListedLicenseById("GPL-2.0-only");
		AnyLicenseInfo apache20 = ListedLicenses.getListedLicenses().getListedLicenseById("Apache-2.0");
		ExtractedLicenseInfo licenseRef1 = new ExtractedLicenseInfo(compareStore, compDocUri, "LicenseRef-1", copyManager, true);
		licenseRef1.setExtractedText(LICENSEREF1_TEXT);
		ExtractedLicenseInfo licenseRef2 = new ExtractedLicenseInfo(compareStore, compDocUri, "LicenseRef-2", copyManager, true);
		licenseRef2.setExtractedText(LICENSEREF2_TEXT);
		ExtractedLicenseInfo licenseRef3 = new ExtractedLicenseInfo(compareStore, compDocUri, "LicenseRef-3", copyManager, true);
		licenseRef3.setExtractedText(LICENSEREF3_TEXT);
		licenseRef3.setName("CyberNeko License");
		licenseRef3.getSeeAlso().add("http://people.apache.org/~andyc/neko/LICENSE");
		licenseRef3.getSeeAlso().add("http://justasample.url.com");
		licenseRef3.setComment("This is tye CyperNeko License");
		ExtractedLicenseInfo licenseRef4 = new ExtractedLicenseInfo(compareStore, compDocUri, "LicenseRef-4", copyManager, true);
		licenseRef4.setExtractedText(LICENSEREF4_TEXT);
		ExtractedLicenseInfo licenseRefBeerware = new ExtractedLicenseInfo(compareStore, compDocUri, "LicenseRef-Beerware-4.2", copyManager, true);
		licenseRefBeerware.setExtractedText(LICENSEREF_BEER_TEXT);
		compareDocument.addExtractedLicenseInfos(licenseRef1);
		compareDocument.addExtractedLicenseInfos(licenseRef2);
		compareDocument.addExtractedLicenseInfos(licenseRef3);
		compareDocument.addExtractedLicenseInfos(licenseRef4);
		compareDocument.addExtractedLicenseInfos(licenseRefBeerware);
		
		ReferenceType locationRefAcmeForge = new ReferenceType(compDocUri + "#" + "LocationRef-acmeforge");
		comparePackages.put("SPDXRef-fromDoap-1", compareDocument.createPackage("SPDXRef-fromDoap-1", 
				"Apache Commons Lang", noAssertionLicense, "NOASSERTION", noAssertionLicense)
				.setHomepage("http://commons.apache.org/proper/commons-lang/")
				.setDownloadLocation("NOASSERTION")
				.setFilesAnalyzed(false)
				.build());
		SpdxPackage spdxrefPackage = compareDocument.createPackage("SPDXRef-Package", 
				"glibc", LicenseInfoFactory.parseSPDXLicenseString("(LicenseRef-3 OR LGPL-2.0-only)", compareStore, compDocUri, copyManager), 
				"Copyright 2008-2010 John Smith", 
				LicenseInfoFactory.parseSPDXLicenseString("(LicenseRef-3 AND LGPL-2.0-only)", compareStore, compDocUri, copyManager))
				.setVersionInfo("2.11.1")
				.setPackageFileName("glibc-2.11.1.tar.gz")
				.setSupplier("Person: Jane Doe (jane.doe@example.com)")
				.setOriginator("Organization: ExampleCodeInspect (contact@example.com)")
				.setHomepage("http://ftp.gnu.org/gnu/glibc")
				.addChecksum(compareDocument.createChecksum(ChecksumAlgorithm.MD5, "624c1abb3664f4b35547e7c73864ad24"))
				.addChecksum(compareDocument.createChecksum(ChecksumAlgorithm.SHA1, "85ed0817af83a24ad8da68c2b5094de69833983c"))
				.addChecksum(compareDocument.createChecksum(ChecksumAlgorithm.SHA256, "11b6d3ee554eedf79299905a98f9b9a04e498210b59f15094c916c91d150efcd"))
				.setPackageVerificationCode(compareDocument.createPackageVerificationCode("d6a770ba38583ed4bb4525bd96e50461655d2758", Arrays.asList(new String[] {"./package.spdx"})))
				.setSourceInfo("uses glibc-2_11-branch from git://sourceware.org/git/glibc.git.")
				.setLicenseInfosFromFile(Arrays.asList(new AnyLicenseInfo[] {gpl2only, licenseRef1, licenseRef2}))
				.setLicenseComments("The license for this project changed with the release of version x.y.  The version of the project included here post-dates the license change.")
				.setSummary("GNU C library.")
				.setDescription("The GNU C Library defines functions that are specified by the ISO C standard, as well as additional features specific to POSIX and other derivatives of the Unix operating system, and extensions specific to GNU systems.")
				.setAttributionText(Arrays.asList(new String[] {"The GNU C Library is free software.  See the file COPYING.LIB for copying conditions, and LICENSES for notices about a few contributions that require these additional notices to be distributed.  License copyright years may be listed using range notation, e.g., 1996-2015, indicating that every year in the range, inclusive, is a copyrightable year that would otherwise be listed individually."}))
				.setDownloadLocation("http://ftp.gnu.org/gnu/glibc/glibc-ports-2.15.tar.gz")
				.setFilesAnalyzed(true)
				.addExternalRef(compareDocument.createExternalRef(ReferenceCategory.SECURITY, 
						ListedReferenceTypes.getListedReferenceTypes().getListedReferenceTypeByName("cpe23Type"), 
						"cpe:2.3:a:pivotal_software:spring_framework:4.1.0:*:*:*:*:*:*:*", null))
				.addExternalRef(compareDocument.createExternalRef(ReferenceCategory.OTHER,
						locationRefAcmeForge, "acmecorp/acmenator/4.1.3-alpha", "This is the external ref for Acme"))	
				.addAnnotation(compareDocument.createAnnotation("Person: Package Commenter", 
						AnnotationType.OTHER, "2011-01-29T18:30:22Z", "Package level annotation"))
				.build();	
		compareDocument.getDocumentDescribes().add(spdxrefPackage);
		comparePackages.put("SPDXRef-Package", spdxrefPackage);
		comparePackages.put("SPDXRef-fromDoap-0", compareDocument.createPackage("SPDXRef-fromDoap-0", 
				"Jena", noAssertionLicense, "NOASSERTION", noAssertionLicense)
				.setHomepage("http://www.openjena.org/")
				.setDownloadLocation("NOASSERTION")
				.setFilesAnalyzed(false)
				.build());
		SpdxPackage spdxrefSaxon = compareDocument.createPackage("SPDXRef-Saxon", 
				"Saxon", LicenseInfoFactory.parseSPDXLicenseString("MPL-1.0", compareStore, compDocUri, copyManager), 
				"NOASSERTION", 
				LicenseInfoFactory.parseSPDXLicenseString("MPL-1.0", compareStore, compDocUri, copyManager))
				.setLicenseComments("Other versions available for a commercial license")
				.setDescription("The Saxon package is a collection of tools for processing XML documents.")
				.setPackageFileName("saxonB-8.8.zip")
				.setVersionInfo("8.8")
				.setHomepage("http://saxon.sourceforge.net/")
				.setDownloadLocation("https://sourceforge.net/projects/saxon/files/Saxon-B/8.8.0.7/saxonb8-8-0-7j.zip/download")
				.addChecksum(compareDocument.createChecksum(ChecksumAlgorithm.SHA1, "85ed0817af83a24ad8da68c2b5094de69833983c"))
				.setFilesAnalyzed(false)
				.build();
		comparePackages.put("SPDXRef-Saxon", spdxrefSaxon);
		
		SpdxFile spdxrefFile = compareDocument.createSpdxFile("SPDXRef-File", 
				"./package/foo.c", LicenseInfoFactory.parseSPDXLicenseString("(LGPL-2.0-only OR LicenseRef-2)", compareStore, compDocUri, copyManager), 
				Arrays.asList(new AnyLicenseInfo[]{gpl2only, licenseRef2}), "Copyright 2008-2010 John Smith", 
				compareDocument.createChecksum(ChecksumAlgorithm.SHA1, "d6a770ba38583ed4bb4525bd96e50461655d2758"))
				.addFileType(FileType.SOURCE)
				.addChecksum(compareDocument.createChecksum(ChecksumAlgorithm.MD5, "624c1abb3664f4b35547e7c73864ad24"))
				.setLicenseComments("The concluded license was taken from the package level that the file was included in.")
				.setNoticeText(SPDXREF_FILE_NOTICE)
				.addFileContributor("The Regents of the University of California")
				.addFileContributor("Modified by Paul Mundt lethal@linux-sh.org")
				.addFileContributor("IBM Corporation")
				.setComment("The concluded license was taken from the package level that the file was included in.\n"+
							"This information was found in the COPYING.txt file in the xyz directory.")
				.addAnnotation(compareDocument.createAnnotation("Person: File Commenter", 
				AnnotationType.OTHER, "2011-01-29T18:30:22Z", "File level annotation"))
				.build();
		compareDocument.getDocumentDescribes().add(spdxrefFile);
		compareFiles.put("SPDXRef-File", spdxrefFile);
		SpdxFile spdxrefCommonsLang = compareDocument.createSpdxFile("SPDXRef-CommonsLangSrc", 
				"./lib-source/commons-lang3-3.1-sources.jar", apache20, 
				Arrays.asList(new AnyLicenseInfo[]{apache20}), "Copyright 2001-2011 The Apache Software Foundation", 
				compareDocument.createChecksum(ChecksumAlgorithm.SHA1, "c2b4e1c67a2d28fced849ee1bb76e7391b93f125"))
				.addFileType(FileType.ARCHIVE)
				.setNoticeText(SPDXREF_COMMONS_NOTICE)
				.addFileContributor("Apache Software Foundation")
				.setComment("This file is used by Jena")
				.addRelationship(compareDocument.createRelationship(new SpdxNoAssertionElement(), RelationshipType.GENERATED_FROM, null))
				.build();
		spdxrefPackage.addFile(spdxrefCommonsLang);
		compareFiles.put("SPDXRef-CommonsLangSrc", spdxrefCommonsLang);
		SpdxFile spdxrefJenalib = compareDocument.createSpdxFile("SPDXRef-JenaLib", 
				"./lib-source/jena-2.6.3-sources.jar", licenseRef1, 
				Arrays.asList(new AnyLicenseInfo[]{licenseRef1}), "(c) Copyright 2000, 2001, 2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009 Hewlett-Packard Development Company, LP", 
				compareDocument.createChecksum(ChecksumAlgorithm.SHA1, "3ab4e1c67a2d28fced849ee1bb76e7391b93f125"))
				.setLicenseComments("This license is used by Jena")
				.addFileType(FileType.ARCHIVE)
				.addFileContributor("Apache Software Foundation")
				.addFileContributor("Hewlett Packard Inc.")
				.setComment("This file belongs to Jena")
				.addRelationship(compareDocument.createRelationship(spdxrefCommonsLang, RelationshipType.DEPENDS_ON, "This relationship replaced a file dependency property value"))
				.build();
		spdxrefPackage.addFile(spdxrefJenalib);
		compareFiles.put("SPDXRef-JenaLib", spdxrefJenalib);
		SpdxFile spdxrefDoap = compareDocument.createSpdxFile("SPDXRef-DoapSource", 
				"./src/org/spdx/parser/DOAPProject.java", apache20, 
				Arrays.asList(new AnyLicenseInfo[]{apache20}), "Copyright 2010, 2011 Source Auditor Inc.", 
				compareDocument.createChecksum(ChecksumAlgorithm.SHA1, "2fd4e1c67a2d28fced849ee1bb76e7391b93eb12"))
				.addFileType(FileType.SOURCE)
				.addFileContributor("Open Logic Inc.")
				.addFileContributor("Black Duck Software In.c")
				.addFileContributor("Source Auditor Inc.")
				.addFileContributor("SPDX Technical Team Members")
				.addFileContributor("Protecode Inc.")
				.addRelationship(compareDocument.createRelationship(spdxrefCommonsLang, RelationshipType.DEPENDS_ON, "This relationship replaced a file dependency property value"))
				.addRelationship(compareDocument.createRelationship(spdxrefJenalib, RelationshipType.DEPENDS_ON, "This relationship replaced a file dependency property value"))
				.build();
		spdxrefPackage.addFile(spdxrefDoap);
		compareFiles.put("SPDXRef-DoapSource", spdxrefDoap);
		
		compareSnippets.put("SPDXRef-Snippet", compareDocument.createSpdxSnippet("SPDXRef-Snippet", 
				"from linux kernel", gpl2only, Arrays.asList(new AnyLicenseInfo[]{gpl2only}), "Copyright 2008-2010 John Smith", 
				spdxrefDoap, 310, 420)
				.setLineRange(5, 23)
				.setLicenseComments("The concluded license was taken from package xyz, from which the snippet was copied into the current file. The concluded license information was found in the COPYING.txt file in package xyz.")
				.setComment("This snippet was identified as significant and highlighted in this Apache-2.0 file, when a commercial scanner identified it as being derived from file foo.c in package xyz which is licensed under GPL-2.0.")
				.build());
		
		compareDocument.addRelationship(compareDocument.createRelationship(spdxrefPackage, RelationshipType.CONTAINS, null));
		compareDocument.createExternalDocumentRef("DocumentRef-spdx-tool-1.2", 
				"http://spdx.org/spdxdocs/spdx-tools-v1.2-3F2504E0-4F89-41D3-9A0C-0305E82C3301", 
				compareDocument.createChecksum(ChecksumAlgorithm.SHA1, "d6a770ba38583ed4bb4525bd96e50461655d2759"));
		ExternalSpdxElement externalElement = new ExternalSpdxElement(compareStore, compDocUri, "DocumentRef-spdx-tool-1.2:SPDXRef-ToolsElement", copyManager, true);
		compareDocument.addRelationship(compareDocument.createRelationship(externalElement, 
				RelationshipType.COPY_OF, null));
		spdxrefFile.addRelationship(compareDocument.createRelationship(spdxrefDoap, 
				RelationshipType.GENERATED_FROM, null));
		spdxrefJenalib.addRelationship(compareDocument.createRelationship(spdxrefPackage, 
				RelationshipType.CONTAINS, null));
		spdxrefPackage.addRelationship(compareDocument.createRelationship(spdxrefJenalib, 
				RelationshipType.CONTAINS, null));
		spdxrefPackage.addRelationship(compareDocument.createRelationship(spdxrefSaxon, 
				RelationshipType.DYNAMIC_LINK, null));
		
		compareDocument.addAnnotation(compareDocument.createAnnotation("Person: Jane Doe ()", 
				AnnotationType.OTHER, "2010-01-29T18:30:22Z", "Document level annotation"));
		compareDocument.addAnnotation(compareDocument.createAnnotation("Person: Joe Reviewer", 
				AnnotationType.REVIEW, "2010-02-10T00:00:00Z", "This is just an example.  Some of the non-standard licenses look like they are actually BSD 3 clause licenses"));
		compareDocument.addAnnotation(compareDocument.createAnnotation("Person: Suzanne Reviewer", 
				AnnotationType.REVIEW, "2011-03-13T00:00:00Z", "Another example reviewer."));
	}

	/* (non-Javadoc)
	 * @see junit.framework.TestCase#tearDown()
	 */
	protected void tearDown() throws Exception {
		super.tearDown();
	}

	/**
	 * Test method for {@link org.spdx.spreadsheetstore.SpreadsheetStore#serialize(java.lang.String, java.io.OutputStream)}.
	 * @throws InvalidSPDXAnalysisException 
	 * @throws IOException 
	 * @throws SpdxCompareException 
	 */
	@SuppressWarnings("unchecked")
    public void testSerialize() throws InvalidSPDXAnalysisException, IOException, SpdxCompareException {
		SpreadsheetStore sst = new SpreadsheetStore(new InMemSpdxStore());
		String documentUri = "http://newdoc/uri";
		ModelCopyManager copyManager = new ModelCopyManager();
		compareStore.getAllItems(compareDocument.getDocumentUri(), SpdxConstants.CLASS_EXTERNAL_DOC_REF).forEach(tv -> {
			try {
				copyManager.copy(sst, documentUri, compareStore, compareDocument.getDocumentUri(), 
						tv.getId(), tv.getType());
			} catch (InvalidSPDXAnalysisException e) {
				throw new RuntimeException(e);
			}
		});
		compareStore.getAllItems(compareDocument.getDocumentUri(), null).forEach(tv -> {
			try {
				if (!SpdxConstants.CLASS_EXTERNAL_DOC_REF.equals(tv.getType())) {
					copyManager.copy(sst, documentUri, compareStore, compareDocument.getDocumentUri(), 
							tv.getId(), tv.getType());
				}
			} catch (InvalidSPDXAnalysisException e) {
				throw new RuntimeException(e);
			}
		});
		
		Path tempFilePath = Files.createTempFile("temp", ".xlsx");
		try {
			try (FileOutputStream out = new FileOutputStream(tempFilePath.toFile())) {
				sst.serialize(documentUri, out);
			}
			SpreadsheetStore resultStore = new SpreadsheetStore(new InMemSpdxStore());
			String resultDocUri;
			try (FileInputStream stream = new FileInputStream(tempFilePath.toFile())) {
				resultDocUri = resultStore.deSerialize(stream, false);
			}
			assertEquals(documentUri, resultDocUri);
			ModelCopyManager cm = new ModelCopyManager();
			SpdxDocument doc = new SpdxDocument(resultStore, resultDocUri, cm, false);
			// Document fields and extracted license infos
			assertDocFields(doc, "SPDX-2.2");
			SpdxComparer comparer = new SpdxComparer();
			comparer.compare(compareDocument, doc);
			assertFalse(comparer.isDifferenceFound());
			// Files
			try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxFile.class)) {
			    elementStream.forEach(element -> {
	                try {
	                    assertTrue(((SpdxElement)element).equivalent(compareFiles.get(((SpdxElement)element).getId())));
	                } catch (InvalidSPDXAnalysisException e) {
	                    fail("Exception: "+e.getMessage());
	                }
	            });
			}
			
			// Packages
           try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxPackage.class)) {
                elementStream.forEach(element -> {
                    try {
                        assertTrue(((SpdxElement)element).equivalent(comparePackages.get(((SpdxElement)element).getId())));
                    } catch (InvalidSPDXAnalysisException e) {
                        fail("Exception: "+e.getMessage());
                    }
                });
            }

			// Snippets
            try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxSnippet.class)) {
               elementStream.forEach(element -> {
                   try {
                       assertTrue(((SpdxElement)element).equivalent(compareSnippets.get(((SpdxElement)element).getId())));
                   } catch (InvalidSPDXAnalysisException e) {
                       fail("Exception: "+e.getMessage());
                   }
               });
            }
		} finally {
			tempFilePath.toFile().delete();
		}
		
	}

	/**
	 * Test method for {@link org.spdx.spreadsheetstore.SpreadsheetStore#deSerialize(java.io.InputStream, boolean)}.
	 * @throws IOException 
	 * @throws InvalidSPDXAnalysisException 
	 */
	@SuppressWarnings("unchecked")
    public void testDeSerialize() throws InvalidSPDXAnalysisException, IOException {
		SpreadsheetStore sst = new SpreadsheetStore(new InMemSpdxStore());
		String documentUri;
		try (FileInputStream stream = new FileInputStream(SPREADSHEET_2_2_FILENAME)) {
			documentUri = sst.deSerialize(stream, false);
		}
		assertEquals("http://spdx.org/spdxdocs/spdx-example-444504E0-4F89-41D3-9A0C-0305E82C3301", documentUri);
		ModelCopyManager cm = new ModelCopyManager();
		SpdxDocument doc = new SpdxDocument(sst, documentUri, cm, false);
		
		// Document fields and extracted license infos
		assertDocFields(doc, "SPDX-2.2");

		// Packages
		try(Stream<SpdxElement> packageStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxPackage.class)) {
		    packageStream.forEach(element -> {
    			try {
    				assertTrue(((SpdxElement)element).equivalent(comparePackages.get(((SpdxElement)element).getId())));
    			} catch (InvalidSPDXAnalysisException e) {
    				fail("Exception: "+e.getMessage());
    			}
    		});
		}
		// Files
        try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxFile.class)) {
          elementStream.forEach(element -> {
                try {
                    assertTrue(((SpdxElement)element).equivalent(compareFiles.get(((SpdxElement)element).getId())));
                } catch (InvalidSPDXAnalysisException e) {
                    fail("Exception: "+e.getMessage());
                }
            });
        }
          

		// Snippets
        try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxSnippet.class)) {
          elementStream.forEach(element -> {
                try {
                    assertTrue(((SpdxElement)element).equivalent(compareSnippets.get(((SpdxElement)element).getId())));
                } catch (InvalidSPDXAnalysisException e) {
                    fail("Exception: "+e.getMessage());
                }
            });
        }
	}
	
	@SuppressWarnings("unchecked")
    public void testDeSerializeXls() throws InvalidSPDXAnalysisException, IOException {
		SpreadsheetStore sst = new SpreadsheetStore(new InMemSpdxStore());
		String documentUri;
		try (FileInputStream stream = new FileInputStream(SPREADSHEET_2_2_FILENAME_XLS)) {
			documentUri = sst.deSerialize(stream, false);
		}
		assertEquals("http://spdx.org/spdxdocs/spdx-example-444504E0-4F89-41D3-9A0C-0305E82C3301", documentUri);
		ModelCopyManager cm = new ModelCopyManager();
		SpdxDocument doc = new SpdxDocument(sst, documentUri, cm, false);
		
		// Document fields and extracted license infos
		assertDocFields(doc, "SPDX-2.2");

		// Packages
		try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxPackage.class)) {
		    elementStream.forEach(element -> {
	            try {
	                assertTrue(((SpdxElement)element).equivalent(comparePackages.get(((SpdxElement)element).getId())));
	            } catch (InvalidSPDXAnalysisException e) {
	                fail("Exception: "+e.getMessage());
	            }
	        });
		}
		
		// Files
        try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxFile.class)) {
            elementStream.forEach(element -> {
                try {
                    assertTrue(((SpdxElement)element).equivalent(compareFiles.get(((SpdxElement)element).getId())));
                } catch (InvalidSPDXAnalysisException e) {
                    fail("Exception: "+e.getMessage());
                }
            });
        }
		
		// Snippets
        try(Stream<SpdxElement> elementStream = (Stream<SpdxElement>)SpdxModelFactory.getElements(sst, documentUri, cm, SpdxSnippet.class)) {
            elementStream.forEach(element -> {
                try {
                    assertTrue(((SpdxElement)element).equivalent(compareSnippets.get(((SpdxElement)element).getId())));
                } catch (InvalidSPDXAnalysisException e) {
                    fail("Exception: "+e.getMessage());
                }
            });
        }
	}
	
	public void testDeSerializeV2() throws InvalidSPDXAnalysisException, IOException {
		SpreadsheetStore sst = new SpreadsheetStore(new InMemSpdxStore());
		String documentUri;
		try (FileInputStream stream = new FileInputStream(SPREADSHEET_2_0_FILENAME)) {
			documentUri = sst.deSerialize(stream, false);
		}
		assertEquals("http://spdx.org/spdxdocs/spdx-example-444504E0-4F89-41D3-9A0C-0305E82C3301", documentUri);
		ModelCopyManager cm = new ModelCopyManager();
		SpdxDocument doc = new SpdxDocument(sst, documentUri, cm, false);
		assertDocFields(doc, "SPDX-2.0");
	}

	private void assertDocFields(SpdxDocument doc, String version) throws InvalidSPDXAnalysisException {
		Collection<Annotation> annotations = doc.getAnnotations();
		assertEquals(3, annotations.size());
		for (Annotation annotation:annotations) {
			if (annotation.getAnnotator().equals("Person: Jane Doe ()")) {
				assertEquals(AnnotationType.OTHER, annotation.getAnnotationType());
				assertEquals("2010-01-29T18:30:22Z", annotation.getAnnotationDate());
				assertEquals("Document level annotation", annotation.getComment());
			} else if (annotation.getAnnotator().equals("Person: Joe Reviewer")) {
				assertEquals(AnnotationType.REVIEW, annotation.getAnnotationType());
				assertEquals("2010-02-10T00:00:00Z", annotation.getAnnotationDate());
				assertEquals("This is just an example.  Some of the non-standard licenses look like they are actually BSD 3 clause licenses", annotation.getComment());
			} else if (annotation.getAnnotator().equals("Person: Suzanne Reviewer")) {
				assertEquals(AnnotationType.REVIEW, annotation.getAnnotationType());
				assertEquals("2011-03-13T00:00:00Z", annotation.getAnnotationDate());
				assertEquals("Another example reviewer.", annotation.getComment());
			} else {
				fail("unexpected annotator: "+annotation.getAnnotator());
			}
		}
		assertEquals("This document was created using SPDX 2.0 using licenses from the web site.",
				doc.getComment().get());
		SpdxCreatorInformation ci = doc.getCreationInfo();
		assertEquals("This package has been shipped in source and binary form.\nThe binaries were created with gcc 4.5.1 and expect to link to\ncompatible system run time libraries.",ci.getComment().get());
		assertTrue(ci.getCreated().startsWith("2010-01-29T"));
		Collection<String> creators = ci.getCreators();
		assertEquals(3, creators.size());
		for (String creator:creators) {
			assertTrue("Tool: LicenseFind-1.0".equals(creator) || 
					"Organization: ExampleCodeInspect ()".equals(creator) || 
					"Person: Jane Doe ()".equals(creator));
		}
		assertEquals("CC0-1.0", doc.getDataLicense().getId());
		Collection<SpdxElement> docDescribes = doc.getDocumentDescribes();
		assertEquals(2, docDescribes.size());
		for (SpdxElement el:docDescribes) {
			assertTrue(el.getId().equals("SPDXRef-File") || el.getId().equals("SPDXRef-Package"));
		}
		Collection<ExternalDocumentRef> externalDocRefs = doc.getExternalDocumentRefs();
		assertEquals(1, externalDocRefs.size());
		for (ExternalDocumentRef edr:externalDocRefs) {
			assertTrue(checksumsMatch(edr.getChecksum(), ChecksumAlgorithm.SHA1, "d6a770ba38583ed4bb4525bd96e50461655d2759"));
			assertEquals("http://spdx.org/spdxdocs/spdx-tools-v1.2-3F2504E0-4F89-41D3-9A0C-0305E82C3301", edr.getSpdxDocumentNamespace());
			assertEquals("DocumentRef-spdx-tool-1.2", edr.getId());
		}
		Collection<ExtractedLicenseInfo> elis = doc.getExtractedLicenseInfos();
		assertEquals(5, elis.size());
		for (ExtractedLicenseInfo eli:elis) {
			assertTrue(eli.getId().equals("LicenseRef-1") || eli.getId().equals("LicenseRef-2") ||
					eli.getId().equals("LicenseRef-3") || eli.getId().equals("LicenseRef-4") ||
					eli.getId().equals("LicenseRef-Beerware-4.2"));
			if (eli.getId().equals("LicenseRef-3")) {
				assertTrue(eli.getExtractedText().startsWith("The CyberNeko Software License, Version 1.0"));
				assertEquals("CyberNeko License", eli.getName());
				Collection<String> seeAlsos = eli.getSeeAlso();
				assertEquals(2, seeAlsos.size());
				for (String sa:seeAlsos) {
					assertTrue(sa.equals("http://people.apache.org/~andyc/neko/LICENSE") ||
							sa.equals("http://justasample.url.com"));
				}
				assertEquals("This is tye CyperNeko License", eli.getComment());
			}
		}
		assertEquals("SPDX-Tools-v2.0", doc.getName().get());
		Collection<Relationship> relationships = doc.getRelationships();
		assertEquals(4, relationships.size());
		for (Relationship rel:relationships) {
			if (rel.getRelatedSpdxElement().get().getId().equals("SPDXRef-Package")) {
				assertTrue(rel.getRelationshipType().equals(RelationshipType.CONTAINS) ||
						rel.getRelationshipType().equals(RelationshipType.DESCRIBES));
			} else if (rel.getRelatedSpdxElement().get().getId().equals("SPDXRef-File")) {
				assertEquals(RelationshipType.DESCRIBES, rel.getRelationshipType());
			} else if (rel.getRelatedSpdxElement().get() instanceof ExternalSpdxElement) {
				assertEquals(RelationshipType.COPY_OF, rel.getRelationshipType());
				assertEquals("DocumentRef-spdx-tool-1.2:SPDXRef-ToolsElement", rel.getRelatedSpdxElement().get().getId());
			} else {
				fail("Unexpected relationship type for related ID "+rel.getRelatedSpdxElement().get().getId());
			}
		}
		assertEquals(version, doc.getSpecVersion());
	}

	private boolean checksumsMatch(Optional<Checksum> checksum, ChecksumAlgorithm algorithm, String value) throws InvalidSPDXAnalysisException {
		if (!checksum.isPresent()) {
			return false;
		}
		return checksum.get().getAlgorithm().equals(algorithm) && checksum.get().getValue().equals(value);
	}

}
