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

import ballerina/io;
import ballerina/test;

const string TEST_DATA_DIR = "tests/resources/testdata/";

// Test record type for parsing
type Employee record {|
    string name;
    int age;
    string department;
|};

@test:Config {
    groups: ["parse"]
}
function testParseBytesToStringArray() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "simple.xlsx");
    string[][] rows = check parseBytes(xlsxData);

    test:assertTrue(rows.length() > 0, "Should have at least one row");
    test:assertTrue(rows[0].length() > 0, "First row should have columns");
}

@test:Config {
    groups: ["parse"]
}
function testParseBytesToRecords() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "employees.xlsx");
    Employee[] employees = check parseBytes(xlsxData);

    test:assertTrue(employees.length() > 0, "Should have at least one employee");
    test:assertEquals(employees[0].name, "John Doe", "First employee name should match");
}

@test:Config {
    groups: ["parse"]
}
function testParseWithHeaderRowOption() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "complex_headers.xlsx");
    ParseOptions opts = {
        headerRow: 2,  // Skip logo/title rows
        dataStartRow: 3
    };
    string[][] rows = check parseBytes(xlsxData, opts);

    test:assertTrue(rows.length() > 0, "Should have data rows");
}

@test:Config {
    groups: ["parse"]
}
function testParseWithSheetSelection() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "multi_sheet.xlsx");
    ParseOptions opts = {
        sheetName: "Sheet2"
    };
    string[][] rows = check parseBytes(xlsxData, opts);

    test:assertTrue(rows.length() > 0, "Should have data from Sheet2");
}

@test:Config {
    groups: ["parse"]
}
function testParseWithSheetIndex() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "multi_sheet.xlsx");
    ParseOptions opts = {
        sheetIndex: 1  // Second sheet (0-based)
    };
    string[][] rows = check parseBytes(xlsxData, opts);

    test:assertTrue(rows.length() > 0, "Should have data from second sheet");
}

@test:Config {
    groups: ["parse", "negative"]
}
function testParseSheetNotFound() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "simple.xlsx");
    ParseOptions opts = {
        sheetName: "NonExistentSheet"
    };
    string[][]|Error result = parseBytes(xlsxData, opts);

    test:assertTrue(result is Error, "Should return error for non-existent sheet");
}

@test:Config {
    groups: ["parse"]
}
function testParseFormulaCachedMode() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "formulas.xlsx");
    ParseOptions opts = {
        formulaMode: CACHED
    };
    string[][] rows = check parseBytes(xlsxData, opts);

    // In CACHED mode, formula cells should return calculated values
    test:assertTrue(rows.length() > 0, "Should have rows");
}

@test:Config {
    groups: ["parse"]
}
function testParseFormulaTextMode() returns error? {
    byte[] xlsxData = check io:fileReadBytes(TEST_DATA_DIR + "formulas.xlsx");
    ParseOptions opts = {
        formulaMode: TEXT
    };
    string[][] rows = check parseBytes(xlsxData, opts);

    // In TEXT mode, formula cells should return formula string
    test:assertTrue(rows.length() > 0, "Should have rows");
    // Check that formula is returned as text (starts with =)
    boolean hasFormula = false;
    foreach string[] row in rows {
        foreach string cell in row {
            if cell.startsWith("=") {
                hasFormula = true;
                break;
            }
        }
    }
    test:assertTrue(hasFormula, "Should have formula text in TEXT mode");
}
