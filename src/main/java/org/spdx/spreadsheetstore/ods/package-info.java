/*
 * SPDX-FileContributor: Arthit Suriyawongkul
 * SPDX-FileCopyrightText: 2026 SPDX Contributors
 * SPDX-FileType: SOURCE
 * SPDX-License-Identifier: Apache-2.0
 */

/**
 * Custom Apache POI adapter layer using the SODS library to provide
 * OpenDocument Spreadsheet (ODS) support.
 * <p>
 * This package maps the Apache POI {@code org.apache.poi.ss.usermodel}
 * interfaces to the underlying {@code com.github.miachm.sods} library.
 * </p>
 * <p>
 * Standard spreadsheet styling elements (colors, fonts, alignments, text
 * wrapping, and cell borders) are fully supported.
 * </p>
 * 
 * <h2>Limitations</h2>
 * <p>
 * Custom cell data formatting (e.g., specific number formats and custom date
 * patterns) has limited support due to the underlying SODS library. SODS only
 * natively supports plain text ("@") or an ISO date format ("YYYY-MM-DD").
 * </p>
 * <p>
 * Consequently, any target cell format representing a date is mapped to the
 * standard {@code "YYYY-MM-DD"} format, while custom numerical and string
 * format patterns remain unstyled.
 * </p>
 */
package org.spdx.spreadsheetstore.ods;
