/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: Copyright (c) 2026 Source Auditor Inc.
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */
package org.spdx.spreadsheetstore;

import java.util.Arrays;
import java.util.List;
import java.util.TimeZone;
import java.util.function.Supplier;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.spdx.spreadsheetstore.ods.OdsWorkbook;

/**
 * Shared helpers for date/timezone regression tests.
 */
public final class SpreadsheetTestUtils {

	private SpreadsheetTestUtils() {
	}

	public static final List<Supplier<Workbook>> WORKBOOK_FACTORIES = Arrays.asList(
			HSSFWorkbook::new, // .xls
			XSSFWorkbook::new, // .xlsx
			OdsWorkbook::new // .ods
	);

	@FunctionalInterface
	public interface ThrowingRunnable {
		void run() throws Exception;
	}

	/**
	 * Run body with the JVM default TimeZone temporarily set to zoneId,
	 * restoring the original default afterward even if body throws.
	 */
	public static void withDefaultTimeZone(String zoneId, ThrowingRunnable body) throws Exception {
		TimeZone original = TimeZone.getDefault();
		try {
			TimeZone.setDefault(TimeZone.getTimeZone(zoneId));
			body.run();
		} finally {
			TimeZone.setDefault(original);
		}
	}
}
