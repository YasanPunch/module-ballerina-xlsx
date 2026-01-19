// Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
//
// WSO2 LLC. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

# Formula handling mode for cells containing formulas.
public enum FormulaMode {
    # Use the last calculated/cached value (default).
    CACHED,
    # Return the formula string as text (e.g., "=SUM(A1:A10)").
    TEXT
}

# Options for parsing XLSX data.
#
# + sheetName - Name of the sheet to read (default: first sheet)
# + sheetIndex - Index of the sheet to read (0-based, alternative to sheetName)
# + headerRow - Row number containing headers (0-based, default: 0)
# + dataStartRow - Row number where data starts (0-based, default: headerRow + 1)
# + includeEmptyRows - Whether to include empty rows in output (default: false)
# + formulaMode - How to handle formula cells (default: CACHED)
# + nilValue - String value to treat as nil (default: empty string)
# + enableConstraintValidation - Whether to validate type constraints (default: true)
public type ParseOptions record {|
    string sheetName?;
    int sheetIndex?;
    int headerRow = 0;
    int dataStartRow?;
    boolean includeEmptyRows = false;
    FormulaMode formulaMode = CACHED;
    string nilValue?;
    boolean enableConstraintValidation = true;
|};

# Options for writing XLSX data.
#
# + sheetName - Name of the sheet to create (default: "Sheet1")
# + writeHeaders - Whether to write headers from record field names (default: true)
# + startRow - Row number to start writing (0-based, default: 0)
public type WriteOptions record {|
    string sheetName = "Sheet1";
    boolean writeHeaders = true;
    int startRow = 0;
|};

# Options for reading rows from a sheet.
#
# + headerRow - Row number containing headers (0-based, default: 0)
# + dataStartRow - Row number where data starts (0-based, default: headerRow + 1)
# + includeEmptyRows - Whether to include empty rows (default: false)
# + formulaMode - How to handle formula cells (default: CACHED)
# + nilValue - String value to treat as nil
public type RowReadOptions record {|
    int headerRow = 0;
    int dataStartRow?;
    boolean includeEmptyRows = false;
    FormulaMode formulaMode = CACHED;
    string nilValue?;
|};

# Options for writing rows to a sheet.
#
# + writeHeaders - Whether to write headers (default: true)
# + startRow - Row number to start writing (0-based, default: 0)
public type RowWriteOptions record {|
    boolean writeHeaders = true;
    int startRow = 0;
|};
