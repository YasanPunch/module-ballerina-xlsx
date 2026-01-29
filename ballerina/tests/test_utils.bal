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

import ballerina/file;
import ballerina/test;

// Setup test data files before running tests
@test:BeforeSuite
function setupTestData() returns error? {
    // Create simple.xlsx
    string[][] simpleData = [
        ["Name", "Age", "City"],
        ["John", "30", "New York"],
        ["Jane", "25", "Los Angeles"],
        ["Bob", "35", "Chicago"]
    ];
    check write(simpleData, TEST_DATA_DIR + "simple.xlsx");

    // Create employees.xlsx with proper headers matching Employee record
    string[][] employeeData = [
        ["name", "age", "department"],
        ["John Doe", "30", "Engineering"],
        ["Jane Smith", "28", "Marketing"],
        ["Bob Johnson", "35", "Sales"]
    ];
    check write(employeeData, TEST_DATA_DIR + "employees.xlsx");

    // Create multi_sheet.xlsx using Workbook API
    Workbook wb = check new Workbook();

    Sheet sheet1 = check wb.createSheet("Sheet1");
    string[][] sheet1Data = [["A1", "B1"], ["A2", "B2"]];
    check sheet1.putRows(sheet1Data);

    Sheet sheet2 = check wb.createSheet("Sheet2");
    string[][] sheet2Data = [["X1", "Y1"], ["X2", "Y2"]];
    check sheet2.putRows(sheet2Data);

    Sheet sheet3 = check wb.createSheet("Sheet3");
    string[][] sheet3Data = [["P1", "Q1"], ["P2", "Q2"]];
    check sheet3.putRows(sheet3Data);

    check wb.save(TEST_DATA_DIR + "multi_sheet.xlsx");
    check wb.close();

    // Create complex_headers.xlsx (with title/logo rows before actual headers)
    string[][] complexHeaderData = [
        ["Company Report", "", ""],           // Row 0: Title
        ["Generated: 2026-01-14", "", ""],    // Row 1: Metadata
        ["Name", "Value", "Status"],          // Row 2: Actual headers
        ["Item1", "100", "Active"],           // Row 3: Data
        ["Item2", "200", "Inactive"]          // Row 4: Data
    ];
    check write(complexHeaderData, TEST_DATA_DIR + "complex_headers.xlsx");

    // Create formulas.xlsx
    // Note: Since we're creating with write which uses POI, we can create cells with formulas
    // For now, we'll create a simple file - formula testing will need native file
    string[][] formulaData = [
        ["A", "B", "Sum"],
        ["10", "20", "=A2+B2"],
        ["15", "25", "=A3+B3"]
    ];
    check write(formulaData, TEST_DATA_DIR + "formulas.xlsx");
}

// Cleanup test data files after running tests
@test:AfterSuite
function cleanupTestData() returns error? {
    // Remove generated test files
    check file:remove(TEST_DATA_DIR + "simple.xlsx");
    check file:remove(TEST_DATA_DIR + "employees.xlsx");
    check file:remove(TEST_DATA_DIR + "multi_sheet.xlsx");
    check file:remove(TEST_DATA_DIR + "complex_headers.xlsx");
    check file:remove(TEST_DATA_DIR + "formulas.xlsx");
}
