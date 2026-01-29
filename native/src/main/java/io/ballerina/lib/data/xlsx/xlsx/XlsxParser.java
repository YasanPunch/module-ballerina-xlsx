/*
 * Copyright (c) 2026, WSO2 LLC. (http://www.wso2.com).
 *
 * WSO2 LLC. licenses this file to you under the Apache License,
 * Version 2.0 (the "License"); you may not use this file except
 * in compliance with the License.
 * You may obtain a copy of the License at
 *
 *    http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing,
 * software distributed under the License is distributed on an
 * "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
 * KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations
 * under the License.
 */

package io.ballerina.lib.data.xlsx.xlsx;

import io.ballerina.lib.data.xlsx.utils.AnnotationUtils;
import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.lib.data.xlsx.utils.UsedRangeDetector;
import io.ballerina.lib.data.xlsx.utils.XlsxConfig;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.creators.TypeCreator;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.types.ArrayType;
import io.ballerina.runtime.api.types.Field;
import io.ballerina.runtime.api.types.MapType;
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.types.UnionType;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.utils.TypeUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Core XLSX parsing logic using Apache POI.
 *
 * @since 0.1.0
 */
public final class XlsxParser {

    private XlsxParser() {
        // Private constructor to prevent instantiation
    }

    /**
     * Parse XLSX bytes to Ballerina value based on target type.
     *
     * @param data       XLSX file bytes
     * @param sheet      Sheet to read (BString for name, Long for index)
     * @param options    Parse options
     * @param targetType Target Ballerina type descriptor
     * @return Parsed value (array of records or array of arrays)
     */
    public static Object parseBytes(byte[] data, Object sheet, BMap<BString, Object> options,
                                    BTypedesc targetType) {
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try (ByteArrayInputStream bis = new ByteArrayInputStream(data);
             Workbook workbook = WorkbookFactory.create(bis)) {

            Sheet selectedSheet = selectSheet(workbook, sheet);
            return parseSheet(selectedSheet, config, targetType);

        } catch (IOException e) {
            return DiagnosticLog.parseError("Failed to parse XLSX: " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error parsing XLSX file: " + e.getMessage(), e);
        }
    }

    /**
     * Parse XLSX from input stream.
     *
     * @param stream     Input stream containing XLSX data
     * @param sheet      Sheet to read (BString for name, Long for index)
     * @param options    Parse options
     * @param targetType Target Ballerina type descriptor
     * @return Parsed value
     */
    public static Object parseAsStream(InputStream stream, Object sheet, BMap<BString, Object> options,
                                     BTypedesc targetType) {
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try (Workbook workbook = WorkbookFactory.create(stream)) {
            Sheet selectedSheet = selectSheet(workbook, sheet);
            return parseSheet(selectedSheet, config, targetType);
        } catch (IOException e) {
            return DiagnosticLog.parseError("Failed to parse XLSX stream: " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error parsing XLSX stream: " + e.getMessage(), e);
        }
    }

    /**
     * Select sheet from workbook based on sheet identifier.
     *
     * @param workbook The workbook to select from
     * @param sheet    Sheet identifier (BString for name, Long for index)
     * @return Selected sheet
     */
    private static Sheet selectSheet(Workbook workbook, Object sheet) {
        Sheet selectedSheet;

        if (sheet instanceof BString) {
            String sheetName = ((BString) sheet).getValue();
            selectedSheet = workbook.getSheet(sheetName);
            if (selectedSheet == null) {
                throw new RuntimeException(DiagnosticLog.sheetNotFoundError(sheetName).getMessage());
            }
        } else if (sheet instanceof Long) {
            int index = ((Long) sheet).intValue();
            if (index < 0 || index >= workbook.getNumberOfSheets()) {
                throw new RuntimeException("Sheet index " + index + " out of range (0-" +
                        (workbook.getNumberOfSheets() - 1) + ")");
            }
            selectedSheet = workbook.getSheetAt(index);
        } else {
            // Default: first sheet
            selectedSheet = workbook.getSheetAt(0);
        }

        return selectedSheet;
    }

    /**
     * Parse a sheet to the target type.
     */
    private static Object parseSheet(Sheet sheet, XlsxConfig config, BTypedesc targetType) {
        Type describingType = targetType.getDescribingType();
        int typeTag = describingType.getTag();

        // Handle array types
        if (typeTag == TypeTags.ARRAY_TAG) {
            ArrayType arrayType = (ArrayType) describingType;
            Type elementType = arrayType.getElementType();

            // Resolve referenced types (important for module-defined types like `type X record {...}`)
            Type resolvedElementType = TypeUtils.getReferredType(elementType);
            int elementTag = resolvedElementType.getTag();

            // string[][] - raw string array
            if (elementTag == TypeTags.ARRAY_TAG) {
                ArrayType innerArrayType = (ArrayType) resolvedElementType;
                Type innerElementType = TypeUtils.getReferredType(innerArrayType.getElementType());
                if (innerElementType.getTag() == TypeTags.STRING_TAG) {
                    return parseToStringArray(sheet, config);
                }
            }

            // record{}[] - array of records
            if (elementTag == TypeTags.RECORD_TYPE_TAG) {
                return parseToRecordArray(sheet, config, (RecordType) resolvedElementType);
            }

            // map<anydata>[] - array of maps
            if (elementTag == TypeTags.MAP_TAG) {
                return parseToMapArray(sheet, config, (MapType) resolvedElementType);
            }
        }

        // Default: parse as string array
        return parseToStringArray(sheet, config);
    }

    /**
     * Parse sheet to string[][].
     */
    private static BArray parseToStringArray(Sheet sheet, XlsxConfig config) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

        if (usedRange == null) {
            // Empty sheet - return empty array
            ArrayType stringArrayType = TypeCreator.createArrayType(
                    TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING));
            return ValueCreator.createArrayValue(stringArrayType);
        }

        int startRow = config.hasExplicitDataStartRow() ?
                config.getDataStartRow() : usedRange.getFirstRow();
        int endRow = usedRange.getLastRow();
        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        List<BArray> rows = new ArrayList<>();
        ArrayType stringType = TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);

        for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
            Row row = sheet.getRow(rowIdx);

            // Skip empty rows unless configured to include them
            if (!config.isIncludeEmptyRows() && UsedRangeDetector.isRowEmpty(row)) {
                continue;
            }

            BArray rowArray = ValueCreator.createArrayValue(stringType);
            for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
                Cell cell = row != null ? row.getCell(colIdx) : null;
                String value = CellConverter.convertToString(cell, config);
                rowArray.append(StringUtils.fromString(value));
            }
            rows.add(rowArray);
        }

        ArrayType resultType = TypeCreator.createArrayType(stringType);
        BArray result = ValueCreator.createArrayValue(resultType);
        for (BArray row : rows) {
            result.append(row);
        }

        return result;
    }

    /**
     * Parse sheet to record[].
     */
    private static BArray parseToRecordArray(Sheet sheet, XlsxConfig config, RecordType recordType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

        if (usedRange == null) {
            // Empty sheet - return empty array
            ArrayType arrayType = TypeCreator.createArrayType(recordType);
            return ValueCreator.createArrayValue(arrayType);
        }

        // Get header row
        int headerRowIndex = config.getHeaderRow();
        Row headerRow = sheet.getRow(headerRowIndex);

        if (headerRow == null) {
            throw new RuntimeException(DiagnosticLog.parseError(
                    "Header row " + headerRowIndex + " is empty",
                    sheet.getSheetName(), headerRowIndex, null).getMessage());
        }

        // Build header-to-column mapping
        Map<String, Integer> headerMap = buildHeaderMap(headerRow, usedRange);

        // Get field mappings
        Map<String, Field> fields = recordType.getFields();
        Map<Integer, FieldMapping> columnToField = new HashMap<>();

        for (Map.Entry<String, Field> entry : fields.entrySet()) {
            String fieldName = entry.getKey();
            Field field = entry.getValue();

            // Check for @xlsx:Name annotation
            String headerName = AnnotationUtils.getHeaderName(recordType, fieldName);

            Integer colIndex = headerMap.get(headerName);
            if (colIndex != null) {
                columnToField.put(colIndex, new FieldMapping(fieldName, field.getFieldType()));
            }
        }

        // Parse data rows
        int dataStartRow = config.getDataStartRow();
        int endRow = usedRange.getLastRow();

        List<BMap<BString, Object>> records = new ArrayList<>();
        ArrayType arrayType = TypeCreator.createArrayType(recordType);

        for (int rowIdx = dataStartRow; rowIdx <= endRow; rowIdx++) {
            Row row = sheet.getRow(rowIdx);

            // Skip empty rows unless configured
            if (!config.isIncludeEmptyRows() && UsedRangeDetector.isRowEmpty(row)) {
                continue;
            }

            BMap<BString, Object> record = ValueCreator.createRecordValue(recordType);

            for (Map.Entry<Integer, FieldMapping> entry : columnToField.entrySet()) {
                int colIdx = entry.getKey();
                FieldMapping mapping = entry.getValue();

                Cell cell = row != null ? row.getCell(colIdx) : null;
                Object value;
                try {
                    value = CellConverter.convert(cell, mapping.type, config);
                } catch (TypeConversionException e) {
                    String cellAddress = getCellAddress(colIdx, rowIdx);
                    throw new RuntimeException(DiagnosticLog.typeConversionError(
                            e.getMessage(), cellAddress, rowIdx, colIdx).getMessage());
                }

                if (value != null) {
                    record.put(StringUtils.fromString(mapping.fieldName), value);
                } else {
                    // Check if field is optional/nilable
                    if (isNilableType(mapping.type)) {
                        record.put(StringUtils.fromString(mapping.fieldName), null);
                    }
                }
            }

            records.add(record);
        }

        BArray result = ValueCreator.createArrayValue(arrayType);
        for (BMap<BString, Object> record : records) {
            result.append(record);
        }

        return result;
    }

    /**
     * Parse sheet to map<anydata>[].
     */
    private static BArray parseToMapArray(Sheet sheet, XlsxConfig config, MapType mapType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

        if (usedRange == null) {
            ArrayType arrayType = TypeCreator.createArrayType(mapType);
            return ValueCreator.createArrayValue(arrayType);
        }

        // Get header row
        int headerRowIndex = config.getHeaderRow();
        Row headerRow = sheet.getRow(headerRowIndex);

        if (headerRow == null) {
            throw new RuntimeException(DiagnosticLog.parseError(
                    "Header row " + headerRowIndex + " is empty",
                    sheet.getSheetName(), headerRowIndex, null).getMessage());
        }

        // Build header map
        Map<String, Integer> headerMap = buildHeaderMap(headerRow, usedRange);
        Map<Integer, String> columnToHeader = new HashMap<>();
        for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
            columnToHeader.put(entry.getValue(), entry.getKey());
        }

        // Parse data rows
        int dataStartRow = config.getDataStartRow();
        int endRow = usedRange.getLastRow();

        Type constraintType = mapType.getConstrainedType();
        ArrayType arrayType = TypeCreator.createArrayType(mapType);
        List<BMap<BString, Object>> maps = new ArrayList<>();

        for (int rowIdx = dataStartRow; rowIdx <= endRow; rowIdx++) {
            Row row = sheet.getRow(rowIdx);

            if (!config.isIncludeEmptyRows() && UsedRangeDetector.isRowEmpty(row)) {
                continue;
            }

            BMap<BString, Object> map = ValueCreator.createMapValue(mapType);

            for (Map.Entry<Integer, String> entry : columnToHeader.entrySet()) {
                int colIdx = entry.getKey();
                String header = entry.getValue();

                Cell cell = row != null ? row.getCell(colIdx) : null;
                Object value;
                try {
                    value = CellConverter.convert(cell, constraintType, config);
                } catch (TypeConversionException e) {
                    String cellAddress = getCellAddress(colIdx, rowIdx);
                    throw new RuntimeException(DiagnosticLog.typeConversionError(
                            e.getMessage(), cellAddress, rowIdx, colIdx).getMessage());
                }

                if (value != null) {
                    map.put(StringUtils.fromString(header), value);
                }
            }

            maps.add(map);
        }

        BArray result = ValueCreator.createArrayValue(arrayType);
        for (BMap<BString, Object> map : maps) {
            result.append(map);
        }

        return result;
    }

    /**
     * Build a map of header names to column indices.
     */
    private static Map<String, Integer> buildHeaderMap(Row headerRow, CellRangeAddress usedRange) {
        Map<String, Integer> headerMap = new HashMap<>();
        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
            Cell cell = headerRow.getCell(colIdx);
            if (cell != null) {
                String headerValue = cell.getStringCellValue();
                if (headerValue != null && !headerValue.trim().isEmpty()) {
                    headerMap.put(headerValue.trim(), colIdx);
                }
            }
        }

        return headerMap;
    }

    /**
     * Check if a type is nilable (includes nil in union).
     */
    private static boolean isNilableType(Type type) {
        if (type.getTag() == TypeTags.UNION_TAG) {
            UnionType unionType = (UnionType) type;
            for (Type memberType : unionType.getMemberTypes()) {
                if (memberType.getTag() == TypeTags.NULL_TAG) {
                    return true;
                }
            }
        }
        return type.isNilable();
    }

    /**
     * Convert column index and row index to Excel cell address (e.g., "A1", "B5").
     */
    private static String getCellAddress(int colIdx, int rowIdx) {
        StringBuilder colName = new StringBuilder();
        int col = colIdx;
        while (col >= 0) {
            colName.insert(0, (char) ('A' + (col % 26)));
            col = col / 26 - 1;
        }
        return colName.toString() + (rowIdx + 1);
    }

    /**
     * Helper class for field mapping.
     */
    private static class FieldMapping {
        final String fieldName;
        final Type type;

        FieldMapping(String fieldName, Type type) {
            this.fieldName = fieldName;
            this.type = type;
        }
    }
}
