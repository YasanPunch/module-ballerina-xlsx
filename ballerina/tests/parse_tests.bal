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
function testParseToStringArray() returns error? {
    string[][] rows = check parse(TEST_DATA_DIR + "simple.xlsx");

    test:assertTrue(rows.length() > 0, "Should have at least one row");
    test:assertTrue(rows[0].length() > 0, "First row should have columns");
}

@test:Config {
    groups: ["parse"]
}
function testParseToRecords() returns error? {
    Employee[] employees = check parse(TEST_DATA_DIR + "employees.xlsx");

    test:assertTrue(employees.length() > 0, "Should have at least one employee");
    test:assertEquals(employees[0].name, "John Doe", "First employee name should match");
}

@test:Config {
    groups: ["parse"]
}
function testParseWithHeaderRowOption() returns error? {
    ParseOptions opts = {
        headerRow: 2,  // Skip logo/title rows
        dataStartRow: 3
    };
    string[][] rows = check parse(TEST_DATA_DIR + "complex_headers.xlsx", 0, opts);

    test:assertTrue(rows.length() > 0, "Should have data rows");
}

@test:Config {
    groups: ["parse"]
}
function testParseWithSheetSelection() returns error? {
    // Sheet selection is now a direct parameter (string for name)
    string[][] rows = check parse(TEST_DATA_DIR + "multi_sheet.xlsx", "Sheet2");

    test:assertTrue(rows.length() > 0, "Should have data from Sheet2");
}

@test:Config {
    groups: ["parse"]
}
function testParseWithSheetIndex() returns error? {
    // Sheet selection is now a direct parameter (int for index, 0-based)
    string[][] rows = check parse(TEST_DATA_DIR + "multi_sheet.xlsx", 1);

    test:assertTrue(rows.length() > 0, "Should have data from second sheet");
}

@test:Config {
    groups: ["parse", "negative"]
}
function testParseSheetNotFound() returns error? {
    // Sheet selection is now a direct parameter
    string[][]|Error result = parse(TEST_DATA_DIR + "simple.xlsx", "NonExistentSheet");

    test:assertTrue(result is Error, "Should return error for non-existent sheet");
}

@test:Config {
    groups: ["parse"]
}
function testParseFormulaCachedMode() returns error? {
    ParseOptions opts = {
        formulaMode: CACHED
    };
    string[][] rows = check parse(TEST_DATA_DIR + "formulas.xlsx", 0, opts);

    // In CACHED mode, formula cells should return calculated values
    test:assertTrue(rows.length() > 0, "Should have rows");
}

@test:Config {
    groups: ["parse"]
}
function testParseFormulaTextMode() returns error? {
    ParseOptions opts = {
        formulaMode: TEXT
    };
    string[][] rows = check parse(TEST_DATA_DIR + "formulas.xlsx", 0, opts);

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
