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
import io.ballerina.runtime.api.types.RecordType;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.utils.TypeUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BObject;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Native handle for Apache POI Sheet.
 * This class wraps the POI Sheet and provides methods called from Ballerina.
 *
 * @since 0.1.0
 */
public final class SheetHandle {

    private static final String SHEET_NATIVE_KEY = "sheetNative";

    private SheetHandle() {
        // Private constructor to prevent instantiation
    }

    /**
     * Initialize a Ballerina Sheet object with a POI Sheet.
     *
     * @param sheetObj Ballerina Sheet object
     * @param sheet    POI Sheet
     */
    static void initSheet(BObject sheetObj, Sheet sheet) {
        sheetObj.addNativeData(SHEET_NATIVE_KEY, sheet);
    }

    /**
     * Get the sheet name.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Sheet name
     */
    public static BString getName(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        return StringUtils.fromString(sheet.getSheetName());
    }

    /**
     * Get the used range of the sheet in A1 notation.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Used range string (e.g., "A1:D50")
     */
    public static BString getUsedRange(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
        return StringUtils.fromString(UsedRangeDetector.toA1Notation(range));
    }

    /**
     * Get the row count of the used range.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Number of rows with data
     */
    public static long getRowCount(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
        return UsedRangeDetector.getRowCount(range);
    }

    /**
     * Get the column count of the used range.
     *
     * @param sheetObj Ballerina Sheet object
     * @return Number of columns with data
     */
    public static long getColumnCount(BObject sheetObj) {
        Sheet sheet = getSheet(sheetObj);
        CellRangeAddress range = UsedRangeDetector.detectUsedRange(sheet);
        return UsedRangeDetector.getColumnCount(range);
    }

    /**
     * Get rows from the sheet.
     *
     * @param sheetObj   Ballerina Sheet object
     * @param options    Read options
     * @param targetType Target type descriptor
     * @return Array of rows (string[][] or record[])
     */
    public static Object getRows(BObject sheetObj, BMap<BString, Object> options, BTypedesc targetType) {
        Sheet sheet = getSheet(sheetObj);
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try {
            Type describingType = targetType.getDescribingType();
            int typeTag = describingType.getTag();

            if (typeTag == TypeTags.ARRAY_TAG) {
                ArrayType arrayType = (ArrayType) describingType;
                Type elementType = arrayType.getElementType();
                // Resolve referenced types (important for module-defined types)
                Type resolvedElementType = TypeUtils.getReferredType(elementType);
                int elementTag = resolvedElementType.getTag();

                // string[][]
                if (elementTag == TypeTags.ARRAY_TAG) {
                    return getRowsAsStringArray(sheet, config);
                }

                // record[]
                if (elementTag == TypeTags.RECORD_TYPE_TAG) {
                    return getRowsAsRecords(sheet, config, (RecordType) resolvedElementType);
                }
            }

            // Default: string[][]
            return getRowsAsStringArray(sheet, config);

        } catch (Exception e) {
            return DiagnosticLog.error("Error getting rows: " + e.getMessage(), e);
        }
    }

    /**
     * Get a single row from the sheet by index.
     *
     * @param sheetObj   Ballerina Sheet object
     * @param index      Row index (0-based, relative to data start row)
     * @param options    Read options
     * @param targetType Target type descriptor
     * @return Single row (string[] or record{})
     */
    public static Object getRow(BObject sheetObj, long index, BMap<BString, Object> options, BTypedesc targetType) {
        Sheet sheet = getSheet(sheetObj);
        XlsxConfig config = XlsxConfig.fromParseOptions(options);

        try {
            CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

            if (usedRange == null) {
                return DiagnosticLog.error("Sheet is empty, cannot get row at index " + index);
            }

            int dataStartRow = config.getDataStartRow();
            int actualRowIndex = dataStartRow + (int) index;
            int endRow = usedRange.getLastRow();

            if (actualRowIndex < dataStartRow || actualRowIndex > endRow) {
                return DiagnosticLog.error("Row index " + index + " out of range (0-" +
                        (endRow - dataStartRow) + ")");
            }

            Row row = sheet.getRow(actualRowIndex);

            Type describingType = targetType.getDescribingType();
            // Resolve referenced types (important for module-defined types)
            Type resolvedType = TypeUtils.getReferredType(describingType);
            int typeTag = resolvedType.getTag();

            // string[] - single row as string array
            if (typeTag == TypeTags.ARRAY_TAG) {
                ArrayType arrayType = (ArrayType) resolvedType;
                Type elementType = TypeUtils.getReferredType(arrayType.getElementType());
                if (elementType.getTag() == TypeTags.STRING_TAG) {
                    return getRowAsStringArray(row, usedRange, config);
                }
            }

            // record{} - single row as record
            if (typeTag == TypeTags.RECORD_TYPE_TAG) {
                return getRowAsRecord(sheet, row, usedRange, config, (RecordType) resolvedType);
            }

            // Default: string[]
            return getRowAsStringArray(row, usedRange, config);

        } catch (Exception e) {
            return DiagnosticLog.error("Error getting row: " + e.getMessage(), e);
        }
    }

    /**
     * Get a single row as string[].
     */
    private static BArray getRowAsStringArray(Row row, CellRangeAddress usedRange, XlsxConfig config) {
        ArrayType stringType = TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
        BArray rowArray = ValueCreator.createArrayValue(stringType);

        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        for (int colIdx = startCol; colIdx <= endCol; colIdx++) {
            Cell cell = row != null ? row.getCell(colIdx) : null;
            String value = CellConverter.convertToString(cell, config);
            rowArray.append(StringUtils.fromString(value));
        }

        return rowArray;
    }

    /**
     * Get a single row as record.
     */
    private static BMap<BString, Object> getRowAsRecord(Sheet sheet, Row row, CellRangeAddress usedRange,
                                                         XlsxConfig config, RecordType recordType) {
        // Get header row
        int headerRowIndex = config.getHeaderRow();
        Row headerRow = sheet.getRow(headerRowIndex);

        if (headerRow == null) {
            throw new RuntimeException("Header row " + headerRowIndex + " is empty");
        }

        // Build header-to-column mapping
        Map<String, Integer> headerMap = buildHeaderMap(headerRow, usedRange);

        // Get field mappings
        Map<String, Field> fields = recordType.getFields();
        Map<Integer, FieldMapping> columnToField = new HashMap<>();

        for (Map.Entry<String, Field> entry : fields.entrySet()) {
            String fieldName = entry.getKey();
            Field field = entry.getValue();
            String headerName = AnnotationUtils.getHeaderName(recordType, fieldName);

            Integer colIndex = headerMap.get(headerName);
            if (colIndex != null) {
                columnToField.put(colIndex, new FieldMapping(fieldName, field.getFieldType()));
            }
        }

        // Create record from row
        BMap<BString, Object> record = ValueCreator.createRecordValue(recordType);

        for (Map.Entry<Integer, FieldMapping> entry : columnToField.entrySet()) {
            int colIdx = entry.getKey();
            FieldMapping mapping = entry.getValue();

            Cell cell = row != null ? row.getCell(colIdx) : null;
            Object value;
            try {
                value = CellConverter.convert(cell, mapping.type, config);
            } catch (TypeConversionException e) {
                int rowIdx = row != null ? row.getRowNum() : -1;
                String cellAddress = getCellAddress(colIdx, rowIdx);
                throw new RuntimeException(DiagnosticLog.typeConversionError(
                        e.getMessage(), cellAddress, rowIdx, colIdx).getMessage());
            }

            if (value != null) {
                record.put(StringUtils.fromString(mapping.fieldName), value);
            }
        }

        return record;
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
     * Put rows into the sheet.
     *
     * @param sheetObj Ballerina Sheet object
     * @param data     Data to write
     * @param options  Write options
     * @return null on success, error on failure
     */
    public static Object putRows(BObject sheetObj, BArray data, BMap<BString, Object> options) {
        Sheet sheet = getSheet(sheetObj);
        return XlsxWriter.writeToSheet(sheet, data, options);
    }

    /**
     * Get rows as string[][].
     */
    private static BArray getRowsAsStringArray(Sheet sheet, XlsxConfig config) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

        ArrayType stringType = TypeCreator.createArrayType(io.ballerina.runtime.api.types.PredefinedTypes.TYPE_STRING);
        ArrayType resultType = TypeCreator.createArrayType(stringType);

        if (usedRange == null) {
            return ValueCreator.createArrayValue(resultType);
        }

        int startRow = config.hasExplicitDataStartRow() ?
                config.getDataStartRow() : usedRange.getFirstRow();
        int endRow = usedRange.getLastRow();
        int startCol = usedRange.getFirstColumn();
        int endCol = usedRange.getLastColumn();

        List<BArray> rows = new ArrayList<>();

        for (int rowIdx = startRow; rowIdx <= endRow; rowIdx++) {
            Row row = sheet.getRow(rowIdx);

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

        BArray result = ValueCreator.createArrayValue(resultType);
        for (BArray row : rows) {
            result.append(row);
        }

        return result;
    }

    /**
     * Get rows as record[].
     */
    private static BArray getRowsAsRecords(Sheet sheet, XlsxConfig config, RecordType recordType) {
        CellRangeAddress usedRange = UsedRangeDetector.detectUsedRange(sheet);

        ArrayType arrayType = TypeCreator.createArrayType(recordType);

        if (usedRange == null) {
            return ValueCreator.createArrayValue(arrayType);
        }

        // Get header row
        int headerRowIndex = config.getHeaderRow();
        Row headerRow = sheet.getRow(headerRowIndex);

        if (headerRow == null) {
            throw new RuntimeException("Header row " + headerRowIndex + " is empty");
        }

        // Build header-to-column mapping
        Map<String, Integer> headerMap = buildHeaderMap(headerRow, usedRange);

        // Get field mappings
        Map<String, Field> fields = recordType.getFields();
        Map<Integer, FieldMapping> columnToField = new HashMap<>();

        for (Map.Entry<String, Field> entry : fields.entrySet()) {
            String fieldName = entry.getKey();
            Field field = entry.getValue();
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

        for (int rowIdx = dataStartRow; rowIdx <= endRow; rowIdx++) {
            Row row = sheet.getRow(rowIdx);

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
     * Build header-to-column mapping.
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
     * Get the native Sheet from Ballerina object.
     */
    private static Sheet getSheet(BObject sheetObj) {
        return (Sheet) sheetObj.getNativeData(SHEET_NATIVE_KEY);
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
