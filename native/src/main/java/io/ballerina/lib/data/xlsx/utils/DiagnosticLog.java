/*
 * Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
 *
 * WSO2 LLC. licenses this file to you under the Apache License,
 * Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package io.ballerina.lib.data.xlsx.utils;

import io.ballerina.runtime.api.creators.ErrorCreator;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BError;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

import java.text.MessageFormat;
import java.util.Locale;
import java.util.ResourceBundle;

/**
 * Utility class for creating diagnostic errors.
 *
 * @since 0.1.0
 */
public final class DiagnosticLog {

    private static final String ERROR_PREFIX = "error.";
    private static ResourceBundle errorBundle;

    private DiagnosticLog() {
        // Private constructor to prevent instantiation
    }

    static {
        try {
            errorBundle = ResourceBundle.getBundle("xlsx_error", Locale.getDefault());
        } catch (Exception e) {
            // Fallback: errors will use default messages
            errorBundle = null;
        }
    }

    /**
     * Create a generic XLSX error.
     *
     * @param message Error message
     * @return BError
     */
    public static BError error(String message) {
        return ErrorCreator.createError(
                ModuleUtils.getModule(),
                Constants.ERROR_TYPE,
                StringUtils.fromString(message),
                null,
                null
        );
    }

    /**
     * Create a generic XLSX error with cause.
     *
     * @param message Error message
     * @param cause   Cause of the error
     * @return BError
     */
    public static BError error(String message, Throwable cause) {
        BError causeError = cause instanceof BError ? (BError) cause :
                ErrorCreator.createError(StringUtils.fromString(cause.getMessage()));
        return ErrorCreator.createError(
                ModuleUtils.getModule(),
                Constants.ERROR_TYPE,
                StringUtils.fromString(message),
                causeError,
                null
        );
    }

    /**
     * Create a diagnostic error with code and arguments.
     *
     * @param code Error code
     * @param args Message arguments
     * @return BError
     */
    public static BError error(DiagnosticErrorCode code, Object... args) {
        String message = formatMessage(code, args);
        return error(message);
    }

    /**
     * Create a parse error.
     *
     * @param message Error message
     * @return BError
     */
    public static BError parseError(String message) {
        return createTypedError(Constants.PARSE_ERROR_TYPE, message, null);
    }

    /**
     * Create a parse error with details.
     *
     * @param message   Error message
     * @param sheetName Sheet name where error occurred
     * @param row       Row number where error occurred
     * @param column    Column number where error occurred
     * @return BError
     */
    public static BError parseError(String message, String sheetName, Integer row, Integer column) {
        BMap<BString, Object> details = createErrorDetails(sheetName, null, row, column);
        return createTypedError(Constants.PARSE_ERROR_TYPE, message, details);
    }

    /**
     * Create a sheet not found error.
     *
     * @param sheetName Sheet name that was not found
     * @return BError
     */
    public static BError sheetNotFoundError(String sheetName) {
        String message = "Sheet '" + sheetName + "' not found in workbook";
        BMap<BString, Object> details = createErrorDetails(sheetName, null, null, null);
        return createTypedError(Constants.SHEET_NOT_FOUND_ERROR_TYPE, message, details);
    }

    /**
     * Create a type conversion error.
     *
     * @param message     Error message
     * @param cellAddress Cell address where error occurred
     * @param row         Row number
     * @param column      Column number
     * @return BError
     */
    public static BError typeConversionError(String message, String cellAddress, Integer row, Integer column) {
        BMap<BString, Object> details = createErrorDetails(null, cellAddress, row, column);
        return createTypedError(Constants.TYPE_CONVERSION_ERROR_TYPE, message, details);
    }

    /**
     * Create a resource limit error.
     *
     * @param message Error message
     * @return BError
     */
    public static BError resourceLimitError(String message) {
        return createTypedError(Constants.RESOURCE_LIMIT_ERROR_TYPE, message, null);
    }

    private static BError createTypedError(String errorType, String message, BMap<BString, Object> details) {
        return ErrorCreator.createError(
                ModuleUtils.getModule(),
                errorType,
                StringUtils.fromString(message),
                null,
                details
        );
    }

    private static BMap<BString, Object> createErrorDetails(String sheetName, String cellAddress,
                                                            Integer row, Integer column) {
        BMap<BString, Object> details = ValueCreator.createMapValue();

        if (sheetName != null) {
            details.put(StringUtils.fromString("sheetName"), StringUtils.fromString(sheetName));
        }
        if (cellAddress != null) {
            details.put(StringUtils.fromString("cellAddress"), StringUtils.fromString(cellAddress));
        }
        if (row != null) {
            details.put(StringUtils.fromString("rowNumber"), (long) row);
        }
        if (column != null) {
            details.put(StringUtils.fromString("columnNumber"), (long) column);
        }

        return details;
    }

    private static String formatMessage(DiagnosticErrorCode code, Object... args) {
        String pattern = getErrorMessage(code);
        if (args.length > 0) {
            return MessageFormat.format(pattern, args);
        }
        return pattern;
    }

    private static String getErrorMessage(DiagnosticErrorCode code) {
        if (errorBundle != null) {
            try {
                return errorBundle.getString(ERROR_PREFIX + code.getMessageKey());
            } catch (Exception e) {
                // Fall through to default
            }
        }
        // Default message based on error code
        return code.getErrorCode() + ": " + code.getMessageKey().replace('.', ' ');
    }
}
