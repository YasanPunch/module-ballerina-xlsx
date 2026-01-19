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

    byte[] xlsxBytes = check toBytes(data);
    test:assertTrue(xlsxBytes.length() > 0, "Should generate non-empty XLSX bytes");

    // Verify by parsing back
    string[][] parsed = check parseBytes(xlsxBytes);
    test:assertEquals(parsed.length(), 3, "Should have 3 rows");
    test:assertEquals(parsed[0][0], "Name", "First cell should be 'Name'");
}

@test:Config {
    groups: ["write"]
}
function testWriteRecords() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true},
        {name: "Bob", age: 35, active: false}
    ];

    byte[] xlsxBytes = check toBytes(people);
    test:assertTrue(xlsxBytes.length() > 0, "Should generate non-empty XLSX bytes");

    // Verify by parsing back as string array (to check headers)
    string[][] parsed = check parseBytes(xlsxBytes);
    test:assertTrue(parsed.length() >= 2, "Should have header + data rows");
}

@test:Config {
    groups: ["write"]
}
function testWriteWithoutHeaders() returns error? {
    Person[] people = [
        {name: "Alice", age: 28, active: true}
    ];

    WriteOptions opts = {
        writeHeaders: false
    };

    byte[] xlsxBytes = check toBytes(people, opts);
    test:assertTrue(xlsxBytes.length() > 0, "Should generate non-empty XLSX bytes");

    // Parse back - should only have data row, no headers
    string[][] parsed = check parseBytes(xlsxBytes);
    test:assertEquals(parsed.length(), 1, "Should have only 1 row (no headers)");
}

@test:Config {
    groups: ["write"]
}
function testWriteWithCustomSheetName() returns error? {
    string[][] data = [["Data"]];

    WriteOptions opts = {
        sheetName: "MyCustomSheet"
    };

    byte[] xlsxBytes = check toBytes(data, opts);

    // Verify by opening as workbook and checking sheet name
    Workbook wb = check openWorkbook(xlsxBytes);
    string[] sheetNames = wb.getSheetNames();
    test:assertEquals(sheetNames[0], "MyCustomSheet", "Sheet name should match");
    check wb.close();
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

    // Write to XLSX
    byte[] xlsxBytes = check toBytes(original);

    // Read back
    string[][] parsed = check parseBytes(xlsxBytes);

    // Verify
    test:assertEquals(parsed.length(), original.length(), "Row count should match");
    test:assertEquals(parsed[0].length(), original[0].length(), "Column count should match");
    test:assertEquals(parsed[1][0], "Alice", "Data should match");
    test:assertEquals(parsed[1][1], "95", "Data should match");
}
