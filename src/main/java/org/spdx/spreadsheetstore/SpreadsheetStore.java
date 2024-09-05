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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;
import java.util.Optional;
import java.util.TreeMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import javax.annotation.Nullable;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.spdx.core.CoreModelObject;
import org.spdx.core.InvalidSPDXAnalysisException;
import org.spdx.library.LicenseInfoFactory;
import org.spdx.library.ModelCopyManager;
import org.spdx.library.SpdxModelFactory;
import org.spdx.library.model.v2.Annotation;
import org.spdx.library.model.v2.ExternalDocumentRef;
import org.spdx.library.model.v2.ExternalRef;
import org.spdx.library.model.v2.ModelObjectV2;
import org.spdx.library.model.v2.Relationship;
import org.spdx.library.model.v2.SpdxConstantsCompatV2;
import org.spdx.library.model.v2.SpdxCreatorInformation;
import org.spdx.library.model.v2.SpdxDocument;
import org.spdx.library.model.v2.SpdxElement;
import org.spdx.library.model.v2.SpdxFile;
import org.spdx.library.model.v2.SpdxModelFactoryCompatV2;
import org.spdx.library.model.v2.SpdxPackage;
import org.spdx.library.model.v2.SpdxSnippet;
import org.spdx.library.model.v2.enumerations.RelationshipType;
import org.spdx.library.model.v2.license.ExtractedLicenseInfo;
import org.spdx.library.model.v2.license.SpdxListedLicense;
import org.spdx.storage.IModelStore;
import org.spdx.storage.ISerializableModelStore;
import org.spdx.storage.simple.ExtendedSpdxStore;

/**
 * SPDX Java Library store which serializes and deserializes to Microsoft Excel Workbooks
 * 
 * @author Gary O'Neall
 *
 */
public class SpreadsheetStore extends ExtendedSpdxStore implements ISerializableModelStore {
	
	static final Logger logger = LoggerFactory.getLogger(SpreadsheetStore.class);
	
	public enum SpreadsheetFormatType {XLS, XLSX};

	private SpreadsheetFormatType spreadsheetFormat;
	
	private static final ThreadLocal<DateFormat> FORMAT = new ThreadLocal<DateFormat>(){
	    @Override
	    protected DateFormat initialValue() {
	        return new SimpleDateFormat(SpdxConstantsCompatV2.SPDX_DATE_FORMAT);
	    }
	  };

	/**
	 * @param baseStore SPDX model store for deserialization/serialization
	 * @param spreadsheetFormat format type XLS or XLSX
	 */
	public SpreadsheetStore(IModelStore baseStore, SpreadsheetFormatType spreadsheetFormat) {
		super(baseStore);
		this.spreadsheetFormat = spreadsheetFormat;
	}
	
	/**
	 * @param baseStore SPDX model store for deserialization/serialization
	 */
	public SpreadsheetStore(IModelStore baseStore) {
		this(baseStore, SpreadsheetFormatType.XLSX);
	}
	
	@Override
	public void serialize(OutputStream stream) throws InvalidSPDXAnalysisException, IOException {
		serialize(stream, null);
	}
	
	@Override
	public void serialize(OutputStream stream, @Nullable CoreModelObject modelObject) throws InvalidSPDXAnalysisException, IOException {
		ModelCopyManager copyManager = new ModelCopyManager();
		SpdxDocument doc = null;
		if (Objects.nonNull(modelObject)) {
			if (modelObject instanceof SpdxDocument) {
				doc = (SpdxDocument)modelObject;
			} else {
				throw new InvalidSPDXAnalysisException("Can not serialize "+modelObject.getClass().toString()+".  Only SpdxDocument is supported");
			}
		} else {
			@SuppressWarnings("unchecked")
			List<SpdxDocument> allDocs = (List<SpdxDocument>)SpdxModelFactory.getSpdxObjects(this, copyManager, 
					SpdxConstantsCompatV2.CLASS_SPDX_DOCUMENT, null, null)
					.collect(Collectors.toList());
			if (allDocs.size() > 1) {
				throw new InvalidSPDXAnalysisException("Ambiguous document to stream - spreadsheet store should only contain one SPDX document");
			}
			if (allDocs.isEmpty()) {
				throw new InvalidSPDXAnalysisException("No SPDX documents to serlialize");
			}
			doc = allDocs.get(0);
		}
		String documentUri = doc.getDocumentUri();
		SpdxSpreadsheet ss = new SpdxSpreadsheet(this, copyManager, documentUri, spreadsheetFormat);
		ss.getOriginsSheet().addDocument(doc);
		Map<String, Collection<ExternalRef>> externalRefs = new TreeMap<String, Collection<ExternalRef>>();
		Map<String, Collection<Relationship>> allRelationships = new TreeMap<>();
		Map<String, Collection<Annotation>> allAnnotations = new TreeMap<>();
		Map<String, String> fileIdToPackageId = copyPackageInfoToSS(documentUri,
				ss.getPackageInfoSheet(), copyManager, externalRefs, allRelationships, allAnnotations);
		copyExternalRefsToSS(externalRefs, ss.getExternalRefsSheet());
		copyExtractedLicenseInfosToSS(doc.getExtractedLicenseInfos(), ss.getExtractedLicenseInfoSheet());
		copyPerFileInfoToSS(documentUri, copyManager, ss.getPerFileSheet(), fileIdToPackageId, allRelationships, allAnnotations);
		copySnippetInfoToSS(documentUri, copyManager, ss.getSnippetSheet(), allRelationships, allAnnotations);		
		allRelationships.put(doc.getId(), doc.getRelationships());
		allAnnotations.put(doc.getId(), doc.getAnnotations());
		copyRelationshipsToSS(allRelationships, ss.getRelationshipsSheet());
		copyAnnotationsToSS(allAnnotations, ss.getAnnotationsSheet());
		ss.resizeRow();
		ss.write(stream);
	}
	
	/**
	 * Copy the annotations ot the annotationsSheet
	 * @param allAnnotations
	 * @param annotationsSheet
	 * @throws SpreadsheetException 
	 */
	private void copyAnnotationsToSS(Map<String, Collection<Annotation>> allAnnotations,
			AnnotationsSheet annotationsSheet) throws SpreadsheetException {
		for (Entry<String, Collection<Annotation>> entry:allAnnotations.entrySet()) {
			Annotation[] annotations = entry.getValue().toArray(new Annotation[entry.getValue().size()]);
			Arrays.sort(annotations);
			for (Annotation annotation:annotations) {
				annotationsSheet.add(annotation, entry.getKey());
			}
		}
	}

	/**
	 * Copy relationships to the relationshipsSheet
	 * @param allRelationships
	 * @param relationshipsSheet
	 * @throws SpreadsheetException 
	 */
	private void copyRelationshipsToSS(Map<String, Collection<Relationship>> allRelationships,
			RelationshipsSheet relationshipsSheet) throws SpreadsheetException {
		for (Entry<String, Collection<Relationship>> entry:allRelationships.entrySet()) {
			Relationship[] relationships = entry.getValue().toArray(new Relationship[entry.getValue().size()]);
			Arrays.sort(relationships);
			for (Relationship relationship:relationships) {
				relationshipsSheet.add(relationship, entry.getKey());
			}
		}
	}

	/**
	 * Copy the snippet information to the snippetSheet and add relationships and annotations to their respective maps
	 * @param documentUri
	 * @param copyManager
	 * @param snippetSheet
	 * @param allRelationships
	 * @param allAnnotations
	 * @throws InvalidSPDXAnalysisException
	 */
	private void copySnippetInfoToSS(String documentUri, ModelCopyManager copyManager, SnippetSheet snippetSheet,
			Map<String, Collection<Relationship>> allRelationships,
			Map<String, Collection<Annotation>> allAnnotations) throws InvalidSPDXAnalysisException {
	    List<SpdxSnippet> snippets;
	    try(
		  @SuppressWarnings("unchecked")
		  Stream<SpdxSnippet> snippetStream = (Stream<SpdxSnippet>)SpdxModelFactory.getSpdxObjects(
	                this, copyManager, SpdxConstantsCompatV2.CLASS_SPDX_SNIPPET, documentUri, documentUri + "#")) {
		    snippets = snippetStream.collect(Collectors.toList());
		}
		Collections.sort(snippets);
		for (SpdxSnippet snippet:snippets) {
			snippetSheet.add(snippet);
			Collection<Relationship> relationships = snippet.getRelationships();
			if (relationships.size() > 0) {
				allRelationships.put(snippet.getId(), relationships);
			}
			Collection<Annotation> annotations = snippet.getAnnotations();
			if (annotations.size() > 0) {
				allAnnotations.put(snippet.getId(), annotations);
			}
		}
	}

	/**
	 * Copy the file information to the perFileSheet and add relationships and annotations to their respective maps
	 * @param documentUri
	 * @param copyManager
	 * @param perFileSheet
	 * @param fileIdToPackageId
	 * @param allRelationships
	 * @param allAnnotations
	 * @throws InvalidSPDXAnalysisException
	 */
	private void copyPerFileInfoToSS(String documentUri, ModelCopyManager copyManager, PerFileSheet perFileSheet,
			Map<String, String> fileIdToPackageId, Map<String, Collection<Relationship>> allRelationships,
			Map<String, Collection<Annotation>> allAnnotations) throws InvalidSPDXAnalysisException {
	    List<SpdxFile> files;
	    try(
		    @SuppressWarnings("unchecked")
		    Stream<SpdxFile> fileStream = (Stream<SpdxFile>)SpdxModelFactory.getSpdxObjects(
	                this, copyManager, SpdxConstantsCompatV2.CLASS_SPDX_FILE, documentUri, documentUri + "#")) {
		    files = fileStream.collect(Collectors.toList());
		}
		Collections.sort(files);
		for (SpdxFile file:files) {
			perFileSheet.add(file, fileIdToPackageId.get(file.getId()));
			Collection<Relationship> relationships = file.getRelationships();
			if (relationships.size() > 0) {
				allRelationships.put(file.getId(), relationships);
			}
			Collection<Annotation> annotations = file.getAnnotations();
			if (annotations.size() > 0) {
				allAnnotations.put(file.getId(), annotations);
			}
		}
	}

	/**
	 * Copy extractedLicenseInfos to the extracteLicenseInfoSheet
	 * @param extractedLicenseInfos
	 * @param extractedLicenseInfoSheet
	 * @throws InvalidSPDXAnalysisException
	 */
	private void copyExtractedLicenseInfosToSS(Collection<ExtractedLicenseInfo> extractedLicenseInfos,
			ExtractedLicenseInfoSheet extractedLicenseInfoSheet) throws InvalidSPDXAnalysisException {
		ExtractedLicenseInfo[] licenses = extractedLicenseInfos.toArray(new ExtractedLicenseInfo[extractedLicenseInfos.size()]);
		Arrays.sort(licenses, new Comparator<ExtractedLicenseInfo>() {

			@Override
			public int compare(ExtractedLicenseInfo o1, ExtractedLicenseInfo o2) {
				int result = 0;
				try {
					if (o1.getName() != null && !(o1.getName().trim().isEmpty())) {
						if (o2.getName() != null && !(o2.getName().trim().isEmpty())) {
							result = o1.getName().compareToIgnoreCase(o2.getName());
						} else {
							result = 1;
						}
					} else {
						result = -1;
					}
				} catch (InvalidSPDXAnalysisException e) {
					result = 0;
				}
				if (result == 0) {
					result = o1.getLicenseId().compareToIgnoreCase(o2.getLicenseId());
				}
				return result;
			}
			
		});
		for(ExtractedLicenseInfo license:licenses) {
			extractedLicenseInfoSheet.add(license.getLicenseId(), license.getExtractedText(), 
					license.getName(),
					license.getSeeAlso(),
					license.getComment());
		}
	}

	/**
	 * Copy the external references to the externalRefSheet
	 * @param externalRefsMap map of package ID to collection of external refs
	 * @param externalRefSheet
	 * @throws SpreadsheetException
	 */
	private void copyExternalRefsToSS(Map<String, Collection<ExternalRef>> externalRefsMap,
			ExternalRefsSheet externalRefSheet) throws SpreadsheetException {
		for (Entry<String, Collection<ExternalRef>> entry:externalRefsMap.entrySet()) {
			ExternalRef[] externalRefs = entry.getValue().toArray(new ExternalRef[entry.getValue().size()]);
			Arrays.sort(externalRefs);
			for (ExternalRef externalRef:externalRefs) {
				externalRefSheet.add(entry.getKey(), externalRef);
			}
		}
	}

	/**
	 * Copy package information from this store into the packageInfoSheet and update the externalRefs, allRelationships, and allAnnotations with collections from the packages
	 * @param documentUri document URI for the document containing the packages
	 * @param packageInfoSheet
	 * @param copyManager
	 * @param externalRefs output parameters of external references
	 * @param allAnnotations 
	 * @param allRelationships 
	 * @return map of file IDs to package ID of the package containing the file
	 * @throws InvalidSPDXAnalysisException
	 */
	private Map<String, String> copyPackageInfoToSS(String documentUri, PackageInfoSheet packageInfoSheet,
			ModelCopyManager copyManager, Map<String, Collection<ExternalRef>> externalRefs, Map<String, Collection<Relationship>> allRelationships, Map<String, Collection<Annotation>> allAnnotations) throws InvalidSPDXAnalysisException {
		Map<String, String> fileIdToPkgId = new HashMap<>();
		List<SpdxPackage> packages;
		
		try (@SuppressWarnings("unchecked")
		Stream<SpdxPackage> packageStream = (Stream<SpdxPackage>)SpdxModelFactory.getSpdxObjects(
                this, copyManager, SpdxConstantsCompatV2.CLASS_SPDX_PACKAGE, documentUri, documentUri + "#")) {
		    packages = packageStream.collect(Collectors.toList());
		}
		Collections.sort(packages);
		for (SpdxPackage pkg:packages) {
			String pkgId = pkg.getId();
			Collection<SpdxFile> files = pkg.getFiles();
			for (SpdxFile file:files) {
				String fileId = file.getId();
				String pkgIdsForFile = fileIdToPkgId.get(fileId);
				if (pkgIdsForFile == null) {
					pkgIdsForFile = pkgId;
				} else {
					pkgIdsForFile = pkgIdsForFile + ", " + pkgId;
				}
				fileIdToPkgId.put(fileId, pkgIdsForFile);
			}
			Collection<ExternalRef> pkgExternalRefs = pkg.getExternalRefs();
			if (pkgExternalRefs != null && pkgExternalRefs.size() > 0) {
				externalRefs.put(pkgId, pkgExternalRefs);
			}
			packageInfoSheet.add(pkg);
			Collection<Relationship> relationships = pkg.getRelationships();
			if (relationships.size() > 0) {
				allRelationships.put(pkg.getId(), relationships);
			}
			Collection<Annotation> annotations = pkg.getAnnotations();
			if (annotations.size() > 0) {
				allAnnotations.put(pkg.getId(), annotations);
			}
		}
		return fileIdToPkgId;
	}
	
	@Override
	public SpdxDocument deSerialize(InputStream stream, boolean overwrite) throws InvalidSPDXAnalysisException, IOException {
		ModelCopyManager copyManager = new ModelCopyManager();
		SpdxSpreadsheet ss = new SpdxSpreadsheet(stream, this, copyManager);
		if (this.exists(ss.getDocumentUri() + "#" + SpdxConstantsCompatV2.SPDX_DOCUMENT_ID)) {
			if (!overwrite) {
				throw new InvalidSPDXAnalysisException("Document "+ss.getDocumentUri()+" already exists.");
			}
			this.clear();
		}
		SpdxDocument document = new SpdxDocument(this, ss.getDocumentUri(), copyManager, true);
		copyDocumentInfoFromSS(ss.getOriginsSheet(), document, ss.getDocumentUri(), copyManager);
		copyExtractedLicenseInfosFromSS(ss.getExtractedLicenseInfoSheet(), document, ss.getDocumentUri(), copyManager);
		// note - non std licenses must be added first so that the text is available
		Map<String, SpdxPackage> pkgIdToPackage = copyPackageInfoFromSS(ss.getPackageInfoSheet(), 
				ss.getExternalRefsSheet(), document);
		// note - packages need to be added before the files so that the files can be added to the packages
		Map<String, SpdxFile> fileIdToFile = copyPerFileInfoFromSS(ss.getPerFileSheet(), document, pkgIdToPackage);
		// note - files need to be added before snippets
		copyPerSnippetInfoFromSS(ss.getSnippetSheet(), document,  fileIdToFile);
		copyAnnotationInfoFromSS(ss.getAnnotationsSheet(), document);
		Map<String, List<String>> packageContainsFileIds = copyRelationshipInfoFromSS(ss.getRelationshipsSheet(), document);
		// Note - the copy missing file contains should be after copying relationships
		copyAnyMissingFileContains(ss.getPerFileSheet(), pkgIdToPackage, fileIdToFile, packageContainsFileIds);
		return document;
	}

	/**
	 * Add any missing package contains file relationships based on the package ID column in the files
	 * sheet.  Note that these should already have been added as relationships.
	 * @param perFileSheet file spreadsheet
	 * @param pkgIdToPackage map of Package ID to package
	 * @param fileIdToFile map of file ID to file
	 * @param packageContainsFileIds map of package ID to list of file IDs that have a contains relationshp
	 * @throws InvalidSPDXAnalysisException 
	 */
	private void copyAnyMissingFileContains(PerFileSheet perFileSheet,
			Map<String, SpdxPackage> pkgIdToPackage,
			Map<String, SpdxFile> fileIdToFile,
			Map<String, List<String>> packageContainsFileIds) throws InvalidSPDXAnalysisException {
		int firstRow = perFileSheet.getFirstDataRow();
		int numFiles = perFileSheet.getNumDataRows();
		for (int i = 0; i < numFiles; i++) {
			String fileId = perFileSheet.getFileId(firstRow+i);
			List<String> pkgIds = perFileSheet.getPackageIds(firstRow+i);
			for (String pkgId:pkgIds) {
				if (!packageContainsFileIds.containsKey(pkgId) ||
				!packageContainsFileIds.get(pkgId).contains(fileId)) {
					SpdxPackage pkg = pkgIdToPackage.get(pkgId);
					SpdxFile file = fileIdToFile.get(fileId);
					if (pkg != null && file != null) {
						if (pkg.getFiles().contains(file)) {
							logger.debug("Skipping duplicate hasFiles");
						} else {
							pkg.addFile(file);
						}
					} else {
						logger.warn("Can not add file "+file.getName()+" to package "+pkgId);
					}
				}
			}
		}
	}

	/**
	 * Copy the extracted information from the extractedLicenseInfoSheet to document
	 * @param extractedLicenseInfoSheet
	 * @param document
	 * @param documentUri
	 * @param copyManager
	 * @throws InvalidSPDXAnalysisException
	 */
	private void copyExtractedLicenseInfosFromSS(ExtractedLicenseInfoSheet extractedLicenseInfoSheet, SpdxDocument document,
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

	/**
	 * Copy document information from the documentInfoSheet to the document
	 * @param documentInfoSheet
	 * @param document
	 * @param documentUri
	 * @param copyManager
	 * @throws InvalidSPDXAnalysisException
	 */
	private void copyDocumentInfoFromSS(DocumentInfoSheet documentInfoSheet, SpdxDocument document, String documentUri, ModelCopyManager copyManager) throws InvalidSPDXAnalysisException {
		Date createdDate = documentInfoSheet.getCreated();
		String created  = FORMAT.get().format(createdDate);
		List<String> createdBys = documentInfoSheet.getCreatedBy();
		SpdxCreatorInformation creationInfo = document.createCreationInfo(createdBys, created); 
		String creatorComment = documentInfoSheet.getAuthorComments();
		if (Objects.nonNull(creatorComment)) {
			creationInfo.setComment(creatorComment);
		}
		String licenseListVersion = documentInfoSheet.getLicenseListVersion();
		if (Objects.nonNull(licenseListVersion)) {
			creationInfo.setLicenseListVersion(licenseListVersion);
		}
		document.setCreationInfo(creationInfo);
		String specVersion = documentInfoSheet.getSPDXVersion();
		document.setSpecVersion(specVersion);
		String dataLicenseId = documentInfoSheet.getDataLicense();
		if (dataLicenseId == null || dataLicenseId.isEmpty()) {
			dataLicenseId = SpdxConstantsCompatV2.SPDX_DATA_LICENSE_ID;
		}
		SpdxListedLicense dataLicense = null;
		try {
			dataLicense = (SpdxListedLicense)LicenseInfoFactory.parseSPDXLicenseStringCompatV2(dataLicenseId, this, documentUri, copyManager);
		} catch (Exception ex) {
			logger.warn("Unable to parse the provided standard license ID.  Using "+SpdxConstantsCompatV2.SPDX_DATA_LICENSE_ID);
			try {
				dataLicense = (SpdxListedLicense)LicenseInfoFactory.parseSPDXLicenseStringCompatV2(SpdxConstantsCompatV2.SPDX_DATA_LICENSE_ID, this, documentUri, copyManager);
			} catch (Exception e) {
				throw(new InvalidSPDXAnalysisException("Unable to get document license"));
			}
		}
		document.setDataLicense(dataLicense);
		String docComment = documentInfoSheet.getDocumentComment();
		if (docComment != null) {
		    docComment = docComment.trim();
			if (!docComment.isEmpty()) {
				document.setComment(docComment);
			}
		}
		String docName = documentInfoSheet.getDocumentName();
		if (docName != null) {
			document.setName(docName);
		}
		Collection<ExternalDocumentRef> externalRefs = documentInfoSheet.getExternalDocumentRefs();
		if (externalRefs != null) {
			document.setExternalDocumentRefs(externalRefs);
		}
	}
	
	/**
	 * Copy package information from the packageInfoSheet to the document
	 * @param packageInfoSheet
	 * @param externalRefsSheet
	 * @param analysis
	 * @return map of ID's to SPDX packages
	 * @throws SpreadsheetException
	 * @throws InvalidSPDXAnalysisException
	 */
	private Map<String, SpdxPackage> copyPackageInfoFromSS(PackageInfoSheet packageInfoSheet,
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
	 * @return map of file ID's to SpdxFiles
	 * @throws SpreadsheetException
	 * @throws InvalidSPDXAnalysisException
	 */
	private Map<String, SpdxFile> copyPerFileInfoFromSS(PerFileSheet perFileSheet,
			SpdxDocument analysis, Map<String, SpdxPackage> pkgIdToPackage) throws SpreadsheetException, InvalidSPDXAnalysisException {
		int firstRow = perFileSheet.getFirstDataRow();
		int numFiles = perFileSheet.getNumDataRows();
		Map<String, SpdxFile> retval = new HashMap<>();
		for (int i = 0; i < numFiles; i++) {
			SpdxFile file = perFileSheet.getFileInfo(firstRow+i);
			retval.put(file.getId(), file);
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
	private void copyPerSnippetInfoFromSS(SnippetSheet snippetSheet,
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
	 * @return map of package IDs to file IDs which have a contains relationship to the package
	 * @throws SpreadsheetException 
	 * @throws InvalidSPDXAnalysisException 
	 */
	private Map<String, List<String>> copyRelationshipInfoFromSS(
			RelationshipsSheet relationshipsSheet, SpdxDocument analysis) throws SpreadsheetException, InvalidSPDXAnalysisException {
		Map<String, List<String>> retval = new HashMap<>();
		int i = relationshipsSheet.getFirstDataRow();
		Relationship relationship = relationshipsSheet.getRelationship(i);
		String id = relationshipsSheet.getElmementId(i);
		while (Objects.nonNull(relationship) && Objects.nonNull(id)) {
			Optional<ModelObjectV2> mo = SpdxModelFactoryCompatV2.getModelObjectV2(analysis.getModelStore(),
					analysis.getDocumentUri(), id, analysis.getCopyManager());
			if (!mo.isPresent()) {
				throw new SpreadsheetException("Missing SPDX element for relationship: "+id);
			}
			if (!(mo.get() instanceof SpdxElement)) {
				throw new SpreadsheetException("ID for SPDX relationship is not of type SpdxElement: "+id);
			}
			if (mo.get() instanceof SpdxPackage && 
					relationship.getRelationshipType().equals(RelationshipType.CONTAINS) && 
					relationship.getRelatedSpdxElement().isPresent() && 
					relationship.getRelatedSpdxElement().get() instanceof SpdxFile) {
				List<String> fileIds = retval.get(mo.get().getId());
				if (fileIds == null) {
					fileIds = new ArrayList<>();
					retval.put(mo.get().getId(), fileIds);
				}
				fileIds.add(relationship.getRelatedSpdxElement().get().getId());
				// need to check for duplicates since the package hasFiles may already contain this
				if (((SpdxPackage)(mo.get())).getFiles().contains(relationship.getRelatedSpdxElement().get())) {
					logger.debug("Skipping duplicate hasFile relationship");
				} else {
					((SpdxElement)(mo.get())).addRelationship(relationship);
				}
			} else {
				((SpdxElement)(mo.get())).addRelationship(relationship);
			}
			i = i + 1;
			relationship = relationshipsSheet.getRelationship(i);
			id = relationshipsSheet.getElmementId(i);
		}
		return retval;
	}

	/**
	 * Copy the annotation information into the model store
	 * @param annotationsSheet
	 * @param analysis
	 * @throws InvalidSPDXAnalysisException 
	 * @throws SpreadsheetException 
	 */
	private void copyAnnotationInfoFromSS(AnnotationsSheet annotationsSheet,
			SpdxDocument analysis) throws InvalidSPDXAnalysisException, SpreadsheetException {
		int i = annotationsSheet.getFirstDataRow();
		Annotation annotation = annotationsSheet.getAnnotation(i);
		String id = annotationsSheet.getElmementId(i);
		while (annotation != null && id != null) {
			Optional<ModelObjectV2> mo = SpdxModelFactoryCompatV2.getModelObjectV2(analysis.getModelStore(),
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
	
	public void unload() {
	    FORMAT.remove();
	}
}
