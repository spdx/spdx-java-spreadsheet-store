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
import java.io.InputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.spdx.library.InvalidSPDXAnalysisException;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxConstants;
import org.spdx.library.model.Annotation;
import org.spdx.library.model.ExternalDocumentRef;
import org.spdx.library.model.ExternalRef;
import org.spdx.library.model.ModelObject;
import org.spdx.library.model.Relationship;
import org.spdx.library.model.SpdxCreatorInformation;
import org.spdx.library.model.SpdxDocument;
import org.spdx.library.model.SpdxElement;
import org.spdx.library.model.SpdxFile;
import org.spdx.library.model.SpdxModelFactory;
import org.spdx.library.model.SpdxPackage;
import org.spdx.library.model.SpdxSnippet;
import org.spdx.library.model.license.ExtractedLicenseInfo;
import org.spdx.library.model.license.LicenseInfoFactory;
import org.spdx.library.model.license.SpdxListedLicense;
import org.spdx.storage.ISerializableModelStore;
import org.spdx.storage.simple.InMemSpdxStore;

/**
 * SPDX Java Library store which serializes and deserializes to Microsoft Excel Workbooks
 * 
 * @author Gary O'Neall
 *
 */
public class SpreadsheetStore extends InMemSpdxStore implements ISerializableModelStore {
	
	static final Logger logger = LoggerFactory.getLogger(SpreadsheetStore.class);
	
	private static final ThreadLocal<DateFormat> format = new ThreadLocal<DateFormat>(){
	    @Override
	    protected DateFormat initialValue() {
	        return new SimpleDateFormat(SpdxConstants.SPDX_DATE_FORMAT);
	    }
	  };

	@Override
	public void serialize(String documentUri, OutputStream stream) throws InvalidSPDXAnalysisException, IOException {
		// TODO Auto-generated method stub
		
	}

	@Override
	public String deSerialize(InputStream stream, boolean overwrite) throws InvalidSPDXAnalysisException, IOException {
		ModelCopyManager copyManager = new ModelCopyManager();
		SpdxSpreadsheet ss = new SpdxSpreadsheet(stream, this, copyManager);
		if (this.exists(ss.getDocumentUri(), SpdxConstants.SPDX_DOCUMENT_ID)) {
			if (!overwrite) {
				throw new InvalidSPDXAnalysisException("Document "+ss.getDocumentUri()+" already exists.");
			}
			this.clear(ss.getDocumentUri());
		}
		SpdxDocument document = SpdxModelFactory.createSpdxDocument(this, ss.getDocumentUri(), copyManager);
		copyOrigins(ss.getOriginsSheet(), document, ss.getDocumentUri(), copyManager);
		ss.getReviewersSheet().addReviewsToDocAnnotations();
		copyExtractedLicenseInfos(ss.getExtractedLicenseInfoSheet(), document, ss.getDocumentUri(), copyManager);
		// note - non std licenses must be added first so that the text is available
		Map<String, SpdxPackage> pkgIdToPackage = copyPackageInfo(ss.getPackageInfoSheet(), 
				ss.getExternalRefsSheet(), document);
		// note - packages need to be added before the files so that the files can be added to the packages
		Map<String, SpdxFile> fileIdToFile = copyPerFileInfo(ss.getPerFileSheet(), document, pkgIdToPackage);
		// note - files need to be added before snippets
		copyPerSnippetInfo(ss.getSnippetSheet(), document,  fileIdToFile);
		copyAnnotationInfo(ss.getAnnotationsSheet(), document);
		copyRelationshipInfo(ss.getRelationshipsSheet(), document);
		return ss.getDocumentUri();
	}

	private void copyExtractedLicenseInfos(ExtractedLicenseInfoSheet extractedLicenseInfoSheet, SpdxDocument document,
			String documentUri, ModelCopyManager copyManager) throws InvalidSPDXAnalysisException {
		int numNonStdLicenses = extractedLicenseInfoSheet.getNumDataRows();
		int firstRow = extractedLicenseInfoSheet.getFirstDataRow();
		for (int i = 0; i < numNonStdLicenses; i++) {
			ExtractedLicenseInfo licenseInfo = new ExtractedLicenseInfo(this, documentUri, 
					extractedLicenseInfoSheet.getIdentifier(firstRow+i), copyManager, true);
			licenseInfo.setExtractedText(extractedLicenseInfoSheet.getExtractedText(firstRow+i));
			licenseInfo.setName(extractedLicenseInfoSheet.getLicenseName(firstRow+i));
			licenseInfo.setSeeAlso(extractedLicenseInfoSheet.getCrossRefUrls(firstRow+i));
			licenseInfo.setComment(extractedLicenseInfoSheet.getComment(firstRow+i));			
			document.addExtractedLicenseInfos(licenseInfo);
		}
	}

	private void copyOrigins(DocumentInfoSheet originsSheet, SpdxDocument document, String documentUri, ModelCopyManager copyManager) throws InvalidSPDXAnalysisException {
		Date createdDate = originsSheet.getCreated();
		String created  = format.get().format(createdDate);
		List<String> createdBys = originsSheet.getCreatedBy();
		SpdxCreatorInformation creationInfo = document.createCreationInfo(createdBys, created); 
		String creatorComment = originsSheet.getAuthorComments();
		if (Objects.nonNull(creatorComment)) {
			creationInfo.setComment(creatorComment);
		}
		String licenseListVersion = originsSheet.getLicenseListVersion();
		if (Objects.nonNull(licenseListVersion)) {
			creationInfo.setLicenseListVersion(licenseListVersion);
		}
		document.setCreationInfo(creationInfo);
		String specVersion = originsSheet.getSPDXVersion();
		document.setSpecVersion(specVersion);
		String dataLicenseId = originsSheet.getDataLicense();
		if (dataLicenseId == null || dataLicenseId.isEmpty()) {
			dataLicenseId = SpdxConstants.SPDX_DATA_LICENSE_ID;
		}
		SpdxListedLicense dataLicense = null;
		try {
			dataLicense = (SpdxListedLicense)LicenseInfoFactory.parseSPDXLicenseString(dataLicenseId, this, documentUri, copyManager);
		} catch (Exception ex) {
			logger.warn("Unable to parse the provided standard license ID.  Using "+SpdxConstants.SPDX_DATA_LICENSE_ID);
			try {
				dataLicense = (SpdxListedLicense)LicenseInfoFactory.parseSPDXLicenseString(SpdxConstants.SPDX_DATA_LICENSE_ID, this, documentUri, copyManager);
			} catch (Exception e) {
				throw(new InvalidSPDXAnalysisException("Unable to get document license"));
			}
		}
		document.setDataLicense(dataLicense);
		String docComment = originsSheet.getDocumentComment();
		if (docComment != null) {
		    docComment = docComment.trim();
			if (!docComment.isEmpty()) {
				document.setComment(docComment);
			}
		}
		String docName = originsSheet.getDocumentName();
		if (docName != null) {
			document.setName(docName);
		}
		Collection<ExternalDocumentRef> externalRefs = originsSheet.getExternalDocumentRefs();
		if (externalRefs != null) {
			document.setExternalDocumentRefs(externalRefs);
		}
	}
	
	private Map<String, SpdxPackage> copyPackageInfo(PackageInfoSheet packageInfoSheet,
			ExternalRefsSheet externalRefsSheet, SpdxDocument analysis) throws SpreadsheetException, InvalidSPDXAnalysisException {
		List<SpdxPackage> packages = packageInfoSheet.getPackages();
		Map<String, SpdxPackage> pkgIdToPackage = new HashMap<>();
		for (SpdxPackage pkg:packages) {
			for (ExternalRef externalRef:externalRefsSheet.getExternalRefsForPkgid(pkg.getId())) {
				pkg.addExternalRef(externalRef);
			}
			pkgIdToPackage.put(pkg.getId(), pkg);
		}
		return pkgIdToPackage;
	}
	
	/**
	 * Copy the file level information
	 * @param perFileSheet
	 * @param analysis
	 * @param pkgIdToPackage
	 * @return
	 * @throws SpreadsheetException
	 * @throws InvalidSPDXAnalysisException
	 */
	private Map<String, SpdxFile> copyPerFileInfo(PerFileSheet perFileSheet,
			SpdxDocument analysis, Map<String, SpdxPackage> pkgIdToPackage) throws SpreadsheetException, InvalidSPDXAnalysisException {
		int firstRow = perFileSheet.getFirstDataRow();
		int numFiles = perFileSheet.getNumDataRows();
		Map<String, SpdxFile> retval = new HashMap<>();
		for (int i = 0; i < numFiles; i++) {
			SpdxFile file = perFileSheet.getFileInfo(firstRow+i);
			retval.put(file.getId(), file);
			List<String> pkgIds = perFileSheet.getPackageIds(firstRow+i);
			for (String pkgId:pkgIds) {
				SpdxPackage pkg = pkgIdToPackage.get(pkgId);
				if (pkg != null) {
					pkg.addFile(file);
				} else {
					logger.warn("Can not add file "+file.getName()+" to package "+pkgId);
				}
			}
		}
		return retval;
	}
	
	/**
	 * Copy snippet information from the spreadsheet to the analysis document
	 * @param snippetSheet
	 * @param analysis
	 * @param fileIdToFile
	 * @throws InvalidSPDXAnalysisException 
	 * @throws SpreadsheetException 
	 */
	private void copyPerSnippetInfo(SnippetSheet snippetSheet,
			SpdxDocument analysis, Map<String, SpdxFile> fileIdToFile) throws InvalidSPDXAnalysisException, SpreadsheetException {
		int i = snippetSheet.getFirstDataRow();
		SpdxSnippet snippet = snippetSheet.getSnippet(i++);
		while (Objects.nonNull(snippet)) {
			snippet = snippetSheet.getSnippet(i++);
		}
	}
	
	/**
	 * Copy the relationships into the model store
	 * @param relationshipsSheet
	 * @param analysis
	 * @throws SpreadsheetException 
	 * @throws InvalidSPDXAnalysisException 
	 */
	private void copyRelationshipInfo(
			RelationshipsSheet relationshipsSheet, SpdxDocument analysis) throws SpreadsheetException, InvalidSPDXAnalysisException {
		int i = relationshipsSheet.getFirstDataRow();
		Relationship relationship = relationshipsSheet.getRelationship(i);
		String id = relationshipsSheet.getElmementId(i);
		while (Objects.nonNull(relationship) && Objects.nonNull(id)) {
			Optional<ModelObject> mo = SpdxModelFactory.getModelObject(analysis.getModelStore(),
					analysis.getDocumentUri(), id, analysis.getCopyManager());
			if (!mo.isPresent()) {
				throw new SpreadsheetException("Missing SPDX element for relationship: "+id);
			}
			if (!(mo.get() instanceof SpdxElement)) {
				throw new SpreadsheetException("ID for SPDX relationship is not of type SpdxElement: "+id);
			}
			((SpdxElement)(mo.get())).addRelationship(relationship);
			i = i + 1;
			relationship = relationshipsSheet.getRelationship(i);
			id = relationshipsSheet.getElmementId(i);
		}
	}

	/**
	 * Copy the annotation information into the model store
	 * @param annotationsSheet
	 * @param analysis
	 * @throws InvalidSPDXAnalysisException 
	 * @throws SpreadsheetException 
	 */
	private void copyAnnotationInfo(AnnotationsSheet annotationsSheet,
			SpdxDocument analysis) throws InvalidSPDXAnalysisException, SpreadsheetException {
		int i = annotationsSheet.getFirstDataRow();
		Annotation annotation = annotationsSheet.getAnnotation(i);
		String id = annotationsSheet.getElmementId(i);
		while (annotation != null && id != null) {
			Optional<ModelObject> mo = SpdxModelFactory.getModelObject(analysis.getModelStore(),
					analysis.getDocumentUri(), id, analysis.getCopyManager());
			if (!mo.isPresent()) {
				throw new SpreadsheetException("Missing SPDX element for annotation: "+id);
			}
			if (!(mo.get() instanceof SpdxElement)) {
				throw new SpreadsheetException("ID for SPDX annotation is not of type SpdxElement: "+id);
			}
			((SpdxElement)(mo.get())).addAnnotation(annotation);
			i = i + 1;
			annotation = annotationsSheet.getAnnotation(i);
			id = annotationsSheet.getElmementId(i);
		}
		
	}
}
