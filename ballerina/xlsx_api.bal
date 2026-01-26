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

import ballerina/jballerina.java;

// ============================================================================
// PRIMARY API - File-based operations (recommended for most use cases)
// ============================================================================

# Parse an XLSX file into Ballerina values.
#
# This is the recommended way to read XLSX files. It reads the specified
# sheet (first sheet by default) and converts rows to the target type.
#
# Supports parsing to:
# - `string[][]` - Raw string array
# - `record{}[]` - Array of records (with header-to-field mapping)
# - `map<anydata>[]` - Array of maps
#
# ```ballerina
# // Parse first sheet as records
# Employee[] employees = check xlsx:parse("employees.xlsx");
#
# // Parse specific sheet by name
# Employee[] sales = check xlsx:parse("report.xlsx", "Sales");
#
# // Parse specific sheet by index with options
# Employee[] data = check xlsx:parse("report.xlsx", 1, {headerRow: 2});
# ```
#
# + path - Path to the XLSX file
# + sheet - Sheet to read: name (string) or index (int, 0-based). Default: 0 (first sheet)
# + options - Parse options
# + t - Target type descriptor
# + return - Parsed data or error
public isolated function parse(string path, string|int sheet = 0, ParseOptions options = {},
        typedesc<anydata[]> t = <>) returns t|Error = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.Native"
} external;

# Write Ballerina data to an XLSX file.
#
# This is the recommended way to write XLSX files. Creates a single-sheet
# XLSX file from the provided data.
#
# Supports writing from:
# - `string[][]` - Raw string array (first row can be headers)
# - `record{}[]` - Array of records (field names become headers)
# - `map<anydata>[]` - Array of maps (keys become headers)
#
# ```ballerina
# Employee[] employees = [{name: "John", age: 30}];
#
# // Write to file
# check xlsx:write(employees, "output.xlsx");
#
# // Write with options
# check xlsx:write(employees, "report.xlsx", sheetName = "Employees");
# ```
#
# + data - Data to write
# + path - Path to the output XLSX file
# + options - Write options
# + return - Error if write fails
public isolated function write(anydata[] data, string path, *WriteOptions options) returns Error? = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.Native"
} external;

// ============================================================================
// STREAMING API - Deferred to v2
// ============================================================================

# Parse XLSX from a byte stream.
#
# **Note**: Deferred to v2. XLSX format requires SharedStringsTable to be
# loaded first, making true streaming complex.
#
# + dataStream - Stream of byte blocks
# + sheet - Sheet to read: name (string) or index (int, 0-based). Default: 0 (first sheet)
# + options - Parse options
# + t - Target type descriptor
# + return - Parsed data or error
public isolated function parseAsStream(stream<byte[], error?> dataStream, string|int sheet = 0,
        ParseOptions options = {}, typedesc<anydata[]> t = <>) returns t|Error = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.Native"
} external;

// ============================================================================
// WORKBOOK API - For multi-sheet operations
// ============================================================================

# Open an XLSX workbook for multi-sheet access.
#
# Use this when you need to:
# - Access multiple sheets efficiently
# - Create a new workbook with multiple sheets
# - Modify and save workbooks
#
# ```ballerina
# // Open from file path (most common)
# xlsx:Workbook wb = check xlsx:openWorkbook("report.xlsx");
#
# // Create new empty workbook
# xlsx:Workbook wb = check xlsx:openWorkbook();
#
# // Work with sheets
# string[] sheets = wb.getSheetNames();
# xlsx:Sheet sheet = check wb.getSheet("Sales");
# Employee[] data = check sheet.getRows();
#
# // Save and close
# check wb.save("updated.xlsx");
# check wb.close();
# ```
#
# + path - File path (string) or nil to create new workbook
# + return - Workbook instance or error
public isolated function openWorkbook(string? path = ()) returns Workbook|Error {
    Workbook workbook = new;
    if path is string {
        check workbook.initFromPath(path);
    } else {
        check workbook.initNew();
    }
    return workbook;
}
