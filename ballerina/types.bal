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

# Annotation to map a record field to a specific Excel column name.
# Use this when the Excel column header doesn't match the Ballerina field name.
#
# ```ballerina
# type Employee record {
#     @xlsx:Name {value: "First Name"}
#     string firstName;
#     @xlsx:Name {value: "Employee ID"}
#     int id;
# };
# ```
#
# When reading, headers "First Name" and "Employee ID" will map to `firstName` and `id`.
# When writing, field names will produce headers "First Name" and "Employee ID".
public type NameConfig record {|
    # The Excel column header name to map to this field.
    string value;
|};

# Annotation to specify the Excel column name for a record field.
public const annotation NameConfig Name on record field;

# Options for parsing XLSX data.
#
# + headerRow - Row containing column headers/names (0-based index).
#               Set to -1 if the sheet has no headers (first row is data).
#               Example: If headers are in row 1 (second row), set this to 1.
# + dataStartRow - Row where actual data begins (0-based index).
#                  If not specified, defaults to headerRow + 1.
#                  Example: If headerRow=0, data starts at row 1 by default.
# + includeEmptyRows - Whether to include empty rows in output (default: false)
# + formulaMode - How to handle formula cells (default: CACHED)
# + enableConstraintValidation - Whether to validate type constraints (default: true).
#                                **Note**: Not yet implemented. Currently ignored.
# + allowDataProjection - Data projection configuration.
#                         **Note**: Not yet implemented. Currently ignored.
#                         Set to `false` to disable data projection.
#                         When enabled (default `{}`):
#                         - `nilAsOptionalField`: Treat nil values as optional field absence
#                         - `absentAsNilableType`: Allow absent columns for nilable types
public type ParseOptions record {|
    int headerRow = 0;
    int dataStartRow?;
    boolean includeEmptyRows = false;
    FormulaMode formulaMode = CACHED;
    boolean enableConstraintValidation = true;
    record {|
        boolean nilAsOptionalField = false;
        boolean absentAsNilableType = false;
    |}|false allowDataProjection = {};
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
# + headerRow - Row containing column headers/names (0-based index).
#               Set to -1 if the sheet has no headers (first row is data).
# + dataStartRow - Row where actual data begins (0-based index).
#                  If not specified, defaults to headerRow + 1.
# + includeEmptyRows - Whether to include empty rows (default: false)
# + formulaMode - How to handle formula cells (default: CACHED)
# + enableConstraintValidation - Whether to validate type constraints (default: true).
#                                **Note**: Not yet implemented. Currently ignored.
# + allowDataProjection - Data projection configuration (see ParseOptions).
#                         **Note**: Not yet implemented. Currently ignored.
public type RowReadOptions record {|
    int headerRow = 0;
    int dataStartRow?;
    boolean includeEmptyRows = false;
    FormulaMode formulaMode = CACHED;
    boolean enableConstraintValidation = true;
    record {|
        boolean nilAsOptionalField = false;
        boolean absentAsNilableType = false;
    |}|false allowDataProjection = {};
|};

# Options for writing rows to a sheet.
#
# + writeHeaders - Whether to write headers (default: true)
# + startRow - Row number to start writing (0-based, default: 0)
public type RowWriteOptions record {|
    boolean writeHeaders = true;
    int startRow = 0;
|};

# Supported data types for writing to XLSX files.
# - `anydata[][]` - 2D array of any data (raw cell values)
# - `record{}[]` - Array of records (field names become headers)
# - `map<anydata>[]` - Array of maps (keys become headers)
public type WritableData anydata[][]|record {}[]|map<anydata>[];
