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

@test:Config {
    groups: ["workbook"]
}
function testOpenWorkbook() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    test:assertTrue(wb.getSheetCount() > 0, "Should have at least one sheet");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testGetSheetNames() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    string[] names = wb.getSheetNames();
    test:assertTrue(names.length() > 0, "Should have sheet names");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testGetSheetByName() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    string[] names = wb.getSheetNames();
    Sheet sheet = check wb.getSheet(names[0]);
    test:assertEquals(sheet.getName(), names[0], "Sheet name should match");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testGetSheetByIndex() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "multi_sheet.xlsx");

    Sheet sheet = check wb.getSheetByIndex(0);
    string[] names = wb.getSheetNames();
    test:assertEquals(sheet.getName(), names[0], "First sheet name should match");
    check wb.close();
}

@test:Config {
    groups: ["workbook", "negative"]
}
function testGetSheetNotFound() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");

    Sheet|SheetNotFoundError result = wb.getSheet("NonExistentSheet");
    test:assertTrue(result is SheetNotFoundError, "Should return SheetNotFoundError");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testCreateNewWorkbook() returns error? {
    Workbook wb = check new Workbook();

    test:assertEquals(wb.getSheetCount(), 0, "New workbook should have no sheets");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testCreateSheet() returns error? {
    Workbook wb = check new Workbook();

    Sheet sheet = check wb.createSheet("TestSheet");
    test:assertEquals(sheet.getName(), "TestSheet", "Sheet name should match");
    test:assertEquals(wb.getSheetCount(), 1, "Should have one sheet");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testWorkbookSave() returns error? {
    Workbook wb = check new Workbook();
    Sheet sheet = check wb.createSheet("Data");

    string[][] data = [["Name", "Value"], ["Test", "123"]];
    check sheet.putRows(data);

    string tempFile = TEST_DATA_DIR + "temp_workbook_save.xlsx";
    check wb.save(tempFile);
    check wb.close();

    // Verify file was created and has content
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by reading back
    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed[0][0], "Name", "Data should be readable");

    // Cleanup
    check file:remove(tempFile);
}

@test:Config {
    groups: ["workbook"]
}
function testSheetGetRows() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    string[][] rows = check sheet.getRows();
    test:assertTrue(rows.length() > 0, "Should have rows");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetPutRows() returns error? {
    Workbook wb = check new Workbook();
    Sheet sheet = check wb.createSheet("TestSheet");

    string[][] data = [
        ["Col1", "Col2"],
        ["A", "B"],
        ["C", "D"]
    ];

    check sheet.putRows(data);

    // Read back
    RowReadOptions opts = {headerRow: 0, dataStartRow: 0};
    string[][] result = check sheet.getRows(opts);
    test:assertEquals(result.length(), 3, "Should have 3 rows");
    check wb.close();
}

@test:Config {
    groups: ["workbook"]
}
function testSheetMetadata() returns error? {
    Workbook wb = check new Workbook(TEST_DATA_DIR + "simple.xlsx");
    Sheet sheet = check wb.getSheetByIndex(0);

    string usedRange = sheet.getUsedRange();
    int rowCount = sheet.getRowCount();
    int colCount = sheet.getColumnCount();

    test:assertTrue(usedRange.length() > 0, "Should have used range");
    test:assertTrue(rowCount > 0, "Should have row count");
    test:assertTrue(colCount > 0, "Should have column count");
    check wb.close();
}
