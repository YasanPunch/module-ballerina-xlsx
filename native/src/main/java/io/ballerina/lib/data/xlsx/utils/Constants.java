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

import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BString;

/**
 * Constants used throughout the XLSX module.
 *
 * @since 0.1.0
 */
public final class Constants {

    private Constants() {
        // Private constructor to prevent instantiation
    }

    // Module information
    public static final String MODULE_NAME = "xlsx";
    public static final String ORG_NAME = "ballerinax";

    // Parse options field names
    public static final BString HEADER_ROW = StringUtils.fromString("headerRow");
    public static final BString DATA_START_ROW = StringUtils.fromString("dataStartRow");
    public static final BString INCLUDE_EMPTY_ROWS = StringUtils.fromString("includeEmptyRows");
    public static final BString FORMULA_MODE = StringUtils.fromString("formulaMode");
    public static final BString ENABLE_CONSTRAINT_VALIDATION = StringUtils.fromString("enableConstraintValidation");

    // Write options field names
    public static final BString WRITE_SHEET_NAME = StringUtils.fromString("sheetName");
    public static final BString WRITE_HEADERS = StringUtils.fromString("writeHeaders");
    public static final BString START_ROW = StringUtils.fromString("startRow");

    // Formula mode values
    public static final String FORMULA_MODE_CACHED = "CACHED";
    public static final String FORMULA_MODE_TEXT = "TEXT";

    // Ballerina type names
    public static final String WORKBOOK_TYPE = "Workbook";
    public static final String SHEET_TYPE = "Sheet";

    // Default values
    public static final String DEFAULT_SHEET_NAME = "Sheet1";
    public static final int DEFAULT_HEADER_ROW = 0;

    // Error type names
    public static final String ERROR_TYPE = "Error";
    public static final String PARSE_ERROR_TYPE = "ParseError";
    public static final String FILE_NOT_FOUND_ERROR_TYPE = "FileNotFoundError";
    public static final String SHEET_NOT_FOUND_ERROR_TYPE = "SheetNotFoundError";
    public static final String TYPE_CONVERSION_ERROR_TYPE = "TypeConversionError";

    // Limits
    public static final int MAX_FILE_SIZE_MB = 100;
    public static final int MAX_ROWS = 1_048_576;  // Excel max rows
    public static final int MAX_COLUMNS = 16_384;   // Excel max columns
}
