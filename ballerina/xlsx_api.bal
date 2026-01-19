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

# Parse XLSX bytes into Ballerina values.
#
# Supports parsing to:
# - `string[][]` - Raw string array
# - `record{}[]` - Array of records (with header-to-field mapping)
# - `map<anydata>[]` - Array of maps
#
# ```ballerina
# // Parse as string array
# string[][] rows = check xlsx:parseBytes(data);
#
# // Parse as records
# type Employee record {| string name; int age; |};
# Employee[] employees = check xlsx:parseBytes(data);
# ```
#
# + data - XLSX file bytes
# + options - Parse options
# + t - Target type descriptor
# + return - Parsed data or error
public isolated function parseBytes(byte[] data, ParseOptions options = {}, typedesc<anydata[]> t = <>)
        returns t|Error = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.Native"
} external;

# Parse XLSX from a byte stream.
#
# Similar to `parseBytes` but accepts a stream of bytes.
# Useful for parsing files from network streams or large files.
#
# ```ballerina
# stream<byte[], io:Error?> fileStream = check io:fileReadBlocksAsStream("data.xlsx");
# string[][] rows = check xlsx:parseStream(fileStream);
# ```
#
# + dataStream - Stream of byte blocks
# + options - Parse options
# + t - Target type descriptor
# + return - Parsed data or error
public isolated function parseStream(stream<byte[], error?> dataStream, ParseOptions options = {},
        typedesc<anydata[]> t = <>) returns t|Error = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.Native"
} external;

# Convert Ballerina data to XLSX bytes.
#
# Supports converting from:
# - `string[][]` - Raw string array (first row can be headers)
# - `record{}[]` - Array of records (field names become headers)
# - `map<anydata>[]` - Array of maps (keys become headers)
#
# ```ballerina
# // From string array
# string[][] data = [["Name", "Age"], ["John", "30"]];
# byte[] xlsx = check xlsx:toBytes(data);
#
# // From records
# Employee[] employees = [{name: "John", age: 30}];
# byte[] xlsx = check xlsx:toBytes(employees);
# ```
#
# + data - Data to convert
# + options - Write options
# + return - XLSX bytes or error
public isolated function toBytes(anydata[] data, *WriteOptions options) returns byte[]|Error = @java:Method {
    'class: "io.ballerina.lib.data.xlsx.Native"
} external;

# Open an XLSX workbook for multi-sheet access.
#
# Use this when you need to:
# - Access multiple sheets efficiently
# - Create a new workbook with multiple sheets
# - Read workbook metadata
#
# ```ballerina
# // Open existing workbook
# xlsx:Workbook wb = check xlsx:openWorkbook(data);
# string[] sheets = wb.getSheetNames();
# xlsx:Sheet sheet = check wb.getSheet("Sales");
# string[][] rows = check sheet.getRows();
# ```
#
# + data - XLSX file bytes (optional, creates new workbook if not provided)
# + return - Workbook instance or error
public isolated function openWorkbook(byte[]? data = ()) returns Workbook|Error {
    Workbook workbook = new;
    if data is byte[] {
        check workbook.initFromBytes(data);
    } else {
        check workbook.initNew();
    }
    return workbook;
}
