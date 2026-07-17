/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore.ods;

import org.apache.poi.ss.usermodel.DataFormat;
import java.util.ArrayList;
import java.util.List;

/**
 * Adapter for Apache POI {@link DataFormat} for ODS documents.
 * <p>
 * Note: The SODS library has limited support for custom data formats.
 * It natively only supports plain text ("@") and an ISO date format
 * ("YYYY-MM-DD"). As a result, exact POI format patterns are not matched in
 * the ODS output. Any date formats will map to "YYYY-MM-DD" internally,
 * while other formats remain unstyled.
 * </p>
 */
public class OdsDataFormat implements DataFormat {

	private final List<String> formats = new ArrayList<>();

	@Override
	public short getFormat(String format) {
		int index = formats.indexOf(format);
		if (index == -1) {
			index = formats.size();
			formats.add(format);
		}
		return (short) index;
	}

	@Override
	public String getFormat(short index) {
		if (index >= 0 && index < formats.size()) {
			return formats.get(index);
		}
		return null;
	}
}
