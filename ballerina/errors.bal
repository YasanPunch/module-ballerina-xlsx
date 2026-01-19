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

# Represents a generic XLSX module error.
public type Error distinct error<ErrorDetails>;

# Represents an error that occurs during XLSX parsing.
public type ParseError distinct error<ErrorDetails>;

# Represents an error when a requested sheet is not found.
public type SheetNotFoundError distinct error<ErrorDetails>;

# Represents an error during type conversion.
public type TypeConversionError distinct error<ErrorDetails>;

# Represents an error when resource limits are exceeded.
public type ResourceLimitError distinct error<ErrorDetails>;

# Details for XLSX errors.
#
# + sheetName - Name of the sheet where error occurred (if applicable)
# + cellAddress - Cell address where error occurred (if applicable)
# + rowNumber - Row number where error occurred (if applicable)
# + columnNumber - Column number where error occurred (if applicable)
public type ErrorDetails record {|
    string sheetName?;
    string cellAddress?;
    int rowNumber?;
    int columnNumber?;
|};
