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

import org.spdx.core.InvalidSPDXAnalysisException;

/**
 * Exceptions related to SPDX Spreadsheets
 *
 * @author Gary O'Neall
 */
public class SpreadsheetException extends InvalidSPDXAnalysisException {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	/**
	 * Construct a SpreadsheetException with the specified message
	 *
	 * @param message The detail message.
	 */
	public SpreadsheetException(String message) {
		super(message);
	}

	/**
	 * Construct a SpreadsheetException with the specified cause
	 *
	 * @param cause The cause of the exception.
	 */
	public SpreadsheetException(Throwable cause) {
		super(cause);
	}

	/**
	 * Construct a SpreadsheetException with the specified message and cause
	 *
	 * @param message The detail message.
	 * @param cause The cause of the exception.
	 */
	public SpreadsheetException(String message, Throwable cause) {
		super(message, cause);
	}

	/**
	 * Construct a SpreadsheetException with the specified details
	 *
	 * @param message The detail message.
	 * @param cause The cause of the exception.
	 * @param enableSuppression Whether suppression is enabled or disabled.
	 * @param writableStackTrace Whether the stack trace should be writable.
	 */
	public SpreadsheetException(String message, Throwable cause, boolean enableSuppression,
			boolean writableStackTrace) {
		super(message, cause, enableSuppression, writableStackTrace);
	}

}
