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

// Test record type for writing
type Person record {|
    string name;
    int age;
    boolean active;
|};

@test:Config {
    groups: ["write"]
}
function testWriteStringArray() returns error? {
    string[][] data = [
        ["Name", "Age", "City"],
        ["John", "30", "New York"],
        ["Jane", "25", "Los Angeles"]
    ];

    string tempFile = TEST_DATA_DIR + "temp_write_string.xlsx";
    check write(data, tempFile);

    // Verify file was created
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by parsing back
    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 3, "Should have 3 rows");
    test:assertEquals(parsed[0][0], "Name", "First cell should be 'Name'");

    // Cleanup
    check file:remove(tempFile);
}

@test:Config {
    groups: ["write"]
}
function testWriteRecords() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true},
        {name: "Bob", age: 35, active: false}
    ];

    string tempFile = TEST_DATA_DIR + "temp_write_records.xlsx";
    check write(people, tempFile);

    // Verify file was created
    test:assertTrue(check file:test(tempFile, file:EXISTS), "File should exist");

    // Verify by parsing back as string array (to check headers)
    string[][] parsed = check parse(tempFile);
    test:assertTrue(parsed.length() >= 2, "Should have header + data rows");

    // Cleanup
    check file:remove(tempFile);
}

@test:Config {
    groups: ["write"]
}
function testWriteWithoutHeaders() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true}
    ];

    string tempFile = TEST_DATA_DIR + "temp_write_no_headers.xlsx";
    check write(people, tempFile, writeHeaders = false);

    // Parse back - should only have data row, no headers
    string[][] parsed = check parse(tempFile);
    test:assertEquals(parsed.length(), 1, "Should have only 1 row (no headers)");

    // Cleanup
    check file:remove(tempFile);
}

@test:Config {
    groups: ["write"]
}
function testWriteWithCustomSheetName() returns error? {
    string[][] data = [["Data"]];

    string tempFile = TEST_DATA_DIR + "temp_write_sheet_name.xlsx";
    check write(data, tempFile, sheetName = "MyCustomSheet");

    // Verify by opening as workbook and checking sheet name
    Workbook wb = check new Workbook(tempFile);
    string[] sheetNames = wb.getSheetNames();
    test:assertEquals(sheetNames[0], "MyCustomSheet", "Sheet name should match");
    check wb.close();

    // Cleanup
    check file:remove(tempFile);
}

@test:Config {
    groups: ["write"]
}
function testRoundTrip() returns error? {
    // Original data
    string[][] original = [
        ["Name", "Score"],
        ["Alice", "95"],
        ["Bob", "87"],
        ["Charlie", "92"]
    ];

    string tempFile = TEST_DATA_DIR + "temp_roundtrip.xlsx";

    // Write to XLSX
    check write(original, tempFile);

    // Read back
    string[][] parsed = check parse(tempFile);

    // Verify
    test:assertEquals(parsed.length(), original.length(), "Row count should match");
    test:assertEquals(parsed[0].length(), original[0].length(), "Column count should match");
    test:assertEquals(parsed[1][0], "Alice", "Data should match");
    test:assertEquals(parsed[1][1], "95", "Data should match");

    // Cleanup
    check file:remove(tempFile);
}

// Test record type with @xlsx:Name annotations
type AnnotatedEmployee record {|
    @Name {value: "First Name"}
    string firstName;
    @Name {value: "Employee ID"}
    int id;
    @Name {value: "Department Name"}
    string department;
|};

@test:Config {
    groups: ["write", "annotation"]
}
function testAnnotatedRecordWrite() returns error? {
    // Write data using annotated record type
    AnnotatedEmployee[] employees = [
        {firstName: "John", id: 101, department: "Engineering"},
        {firstName: "Jane", id: 102, department: "Marketing"}
    ];

    string tempFile = TEST_DATA_DIR + "temp_annotated.xlsx";
    check write(employees, tempFile);

    // Parse back as string[][] to verify headers match annotation values
    string[][] parsed = check parse(tempFile);

    test:assertEquals(parsed.length(), 3, "Should have header + 2 data rows");

    // Verify headers are annotation values, not field names
    string[] headers = parsed[0];
    test:assertTrue(headers.indexOf("First Name") != (), "Should have 'First Name' header");
    test:assertTrue(headers.indexOf("Employee ID") != (), "Should have 'Employee ID' header");
    test:assertTrue(headers.indexOf("Department Name") != (), "Should have 'Department Name' header");

    // Verify field names are NOT used as headers
    test:assertTrue(headers.indexOf("firstName") == (), "Should NOT have 'firstName' header");
    test:assertTrue(headers.indexOf("id") == (), "Should NOT have 'id' header");

    // Cleanup
    check file:remove(tempFile);
}

@test:Config {
    groups: ["parse", "annotation"]
}
function testAnnotatedRecordParse() returns error? {
    // First write data with annotated record to create test file
    AnnotatedEmployee[] employees = [
        {firstName: "Alice", id: 201, department: "Sales"},
        {firstName: "Bob", id: 202, department: "Support"}
    ];

    string tempFile = TEST_DATA_DIR + "temp_annotated_parse.xlsx";
    check write(employees, tempFile);

    // Now parse back using the same annotated record type
    AnnotatedEmployee[] parsed = check parse(tempFile);

    test:assertEquals(parsed.length(), 2, "Should have 2 employee records");
    test:assertEquals(parsed[0].firstName, "Alice", "First employee firstName should match");
    test:assertEquals(parsed[0].id, 201, "First employee id should match");
    test:assertEquals(parsed[0].department, "Sales", "First employee department should match");
    test:assertEquals(parsed[1].firstName, "Bob", "Second employee firstName should match");

    // Cleanup
    check file:remove(tempFile);
}
