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

import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;

/**
 * Configuration holder for XLSX parsing and writing options.
 *
 * @since 0.1.0
 */
public class XlsxConfig {

    // Parse options
    private int headerRow = Constants.DEFAULT_HEADER_ROW;
    private Integer dataStartRow;
    private boolean includeEmptyRows = false;
    private String formulaMode = Constants.FORMULA_MODE_CACHED;
    private boolean enableConstraintValidation = true;

    // Write options
    private String writeSheetName = Constants.DEFAULT_SHEET_NAME;
    private boolean writeHeaders = true;
    private int startRow = 0;

    /**
     * Create config from Ballerina parse options map.
     *
     * @param options Ballerina options map
     * @return XlsxConfig instance
     */
    public static XlsxConfig fromParseOptions(BMap<BString, Object> options) {
        XlsxConfig config = new XlsxConfig();

        if (options == null) {
            return config;
        }

        // Header configuration
        Object headerRowVal = options.get(Constants.HEADER_ROW);
        if (headerRowVal != null) {
            config.headerRow = ((Long) headerRowVal).intValue();
        }

        Object dataStartRowVal = options.get(Constants.DATA_START_ROW);
        if (dataStartRowVal != null) {
            config.dataStartRow = ((Long) dataStartRowVal).intValue();
        }

        // Row handling
        Object includeEmptyVal = options.get(Constants.INCLUDE_EMPTY_ROWS);
        if (includeEmptyVal != null) {
            config.includeEmptyRows = (Boolean) includeEmptyVal;
        }

        // Formula handling
        Object formulaModeVal = options.get(Constants.FORMULA_MODE);
        if (formulaModeVal != null) {
            config.formulaMode = formulaModeVal.toString();
        }

        // Constraint validation
        Object constraintVal = options.get(Constants.ENABLE_CONSTRAINT_VALIDATION);
        if (constraintVal != null) {
            config.enableConstraintValidation = (Boolean) constraintVal;
        }

        return config;
    }

    /**
     * Create config from Ballerina write options map.
     *
     * @param options Ballerina options map
     * @return XlsxConfig instance
     */
    public static XlsxConfig fromWriteOptions(BMap<BString, Object> options) {
        XlsxConfig config = new XlsxConfig();

        if (options == null) {
            return config;
        }

        Object sheetNameVal = options.get(Constants.WRITE_SHEET_NAME);
        if (sheetNameVal != null) {
            config.writeSheetName = sheetNameVal.toString();
        }

        Object writeHeadersVal = options.get(Constants.WRITE_HEADERS);
        if (writeHeadersVal != null) {
            config.writeHeaders = (Boolean) writeHeadersVal;
        }

        Object startRowVal = options.get(Constants.START_ROW);
        if (startRowVal != null) {
            config.startRow = ((Long) startRowVal).intValue();
        }

        return config;
    }

    // Getters

    public int getHeaderRow() {
        return headerRow;
    }

    public int getDataStartRow() {
        // Default: row after header
        return dataStartRow != null ? dataStartRow : headerRow + 1;
    }

    public boolean hasExplicitDataStartRow() {
        return dataStartRow != null;
    }

    public boolean isIncludeEmptyRows() {
        return includeEmptyRows;
    }

    public String getFormulaMode() {
        return formulaMode;
    }

    public boolean isFormulaModeText() {
        return Constants.FORMULA_MODE_TEXT.equals(formulaMode);
    }

    public boolean isEnableConstraintValidation() {
        return enableConstraintValidation;
    }

    public String getWriteSheetName() {
        return writeSheetName;
    }

    public boolean isWriteHeaders() {
        return writeHeaders;
    }

    public int getStartRow() {
        return startRow;
    }
}
