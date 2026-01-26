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

import io.ballerina.lib.data.xlsx.utils.XlsxConfig;
import io.ballerina.runtime.api.types.TypeTags;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.types.Type;
import io.ballerina.runtime.api.utils.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Utility class for converting Excel cells to Ballerina values.
 *
 * @since 0.1.0
 */
public final class CellConverter {

    private static final SimpleDateFormat DATE_FORMAT = new SimpleDateFormat("yyyy-MM-dd");

    private CellConverter() {
        // Private constructor to prevent instantiation
    }

    /**
     * Convert an Excel cell to a Ballerina value.
     *
     * @param cell       The cell to convert
     * @param targetType The target Ballerina type (may be null for string default)
     * @param config     Parsing configuration
     * @return The converted value, or null for empty cells
     */
    public static Object convert(Cell cell, Type targetType, XlsxConfig config) {
        if (cell == null) {
            return null;
        }

        CellType cellType = cell.getCellType();

        // Handle formula cells
        if (cellType == CellType.FORMULA) {
            return handleFormula(cell, targetType, config);
        }

        return convertByType(cell, cellType, targetType, config);
    }

    /**
     * Convert a cell value to string (for string[][] output).
     *
     * @param cell   The cell to convert
     * @param config Parsing configuration
     * @return String value
     */
    public static String convertToString(Cell cell, XlsxConfig config) {
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCellType();

        // Handle formula cells
        if (cellType == CellType.FORMULA) {
            if (config.isFormulaModeText()) {
                return "=" + cell.getCellFormula();
            }
            // CACHED mode - get cached result
            cellType = cell.getCachedFormulaResultType();
        }

        switch (cellType) {
            case STRING:
                return cell.getStringCellValue();

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    Date date = cell.getDateCellValue();
                    return DATE_FORMAT.format(date);
                }
                double numValue = cell.getNumericCellValue();
                // Format as integer if it's a whole number
                if (numValue == Math.floor(numValue) && !Double.isInfinite(numValue)) {
                    return String.valueOf((long) numValue);
                }
                return String.valueOf(numValue);

            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());

            case BLANK:
                return "";

            case ERROR:
                return "#ERROR";

            default:
                return cell.toString();
        }
    }

    private static Object handleFormula(Cell cell, Type targetType, XlsxConfig config) {
        // TEXT mode - return formula string
        if (config.isFormulaModeText()) {
            return StringUtils.fromString("=" + cell.getCellFormula());
        }

        // CACHED mode - return last calculated value
        CellType cachedType = cell.getCachedFormulaResultType();
        return convertByType(cell, cachedType, targetType, config);
    }

    private static Object convertByType(Cell cell, CellType cellType, Type targetType, XlsxConfig config) {
        switch (cellType) {
            case STRING:
                String strValue = cell.getStringCellValue();
                // Check for nil value (null or empty string)
                if (isNilValue(strValue)) {
                    return null;
                }
                return convertStringToTarget(strValue, targetType);

            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return convertDate(cell.getDateCellValue(), targetType);
                }
                return convertNumeric(cell.getNumericCellValue(), targetType);

            case BOOLEAN:
                return convertBoolean(cell.getBooleanCellValue(), targetType);

            case BLANK:
                return null;

            case ERROR:
                // Return null for error cells in data context
                return null;

            default:
                return StringUtils.fromString(cell.toString());
        }
    }

    private static boolean isNilValue(String value) {
        return value == null || value.trim().isEmpty();
    }

    private static Object convertStringToTarget(String value, Type targetType) {
        if (targetType == null) {
            return StringUtils.fromString(value);
        }

        int typeTag = targetType.getTag();

        switch (typeTag) {
            case TypeTags.INT_TAG:
                try {
                    return Long.parseLong(value.trim());
                } catch (NumberFormatException e) {
                    // Try parsing as double first
                    try {
                        return (long) Double.parseDouble(value.trim());
                    } catch (NumberFormatException e2) {
                        return StringUtils.fromString(value);
                    }
                }

            case TypeTags.FLOAT_TAG:
                try {
                    return Double.parseDouble(value.trim());
                } catch (NumberFormatException e) {
                    return StringUtils.fromString(value);
                }

            case TypeTags.DECIMAL_TAG:
                try {
                    return ValueCreator.createDecimalValue(new BigDecimal(value.trim()));
                } catch (NumberFormatException e) {
                    return StringUtils.fromString(value);
                }

            case TypeTags.BOOLEAN_TAG:
                return parseBoolean(value);

            case TypeTags.STRING_TAG:
            default:
                return StringUtils.fromString(value);
        }
    }

    private static Object convertNumeric(double value, Type targetType) {
        if (targetType == null) {
            // Default: return as int if whole number, else decimal
            if (value == Math.floor(value) && !Double.isInfinite(value)) {
                return (long) value;
            }
            return ValueCreator.createDecimalValue(BigDecimal.valueOf(value));
        }

        int typeTag = targetType.getTag();

        switch (typeTag) {
            case TypeTags.INT_TAG:
                return (long) value;

            case TypeTags.FLOAT_TAG:
                return value;

            case TypeTags.DECIMAL_TAG:
                return ValueCreator.createDecimalValue(BigDecimal.valueOf(value));

            case TypeTags.STRING_TAG:
                if (value == Math.floor(value) && !Double.isInfinite(value)) {
                    return StringUtils.fromString(String.valueOf((long) value));
                }
                return StringUtils.fromString(String.valueOf(value));

            case TypeTags.BOOLEAN_TAG:
                return value != 0;

            default:
                return ValueCreator.createDecimalValue(BigDecimal.valueOf(value));
        }
    }

    private static Object convertBoolean(boolean value, Type targetType) {
        if (targetType == null) {
            return value;
        }

        int typeTag = targetType.getTag();

        switch (typeTag) {
            case TypeTags.BOOLEAN_TAG:
                return value;

            case TypeTags.STRING_TAG:
                return StringUtils.fromString(String.valueOf(value));

            case TypeTags.INT_TAG:
                return value ? 1L : 0L;

            case TypeTags.FLOAT_TAG:
                return value ? 1.0 : 0.0;

            default:
                return value;
        }
    }

    private static Object convertDate(Date date, Type targetType) {
        // For now, convert dates to string in ISO format
        // Future: support time:Date type
        String dateStr = DATE_FORMAT.format(date);
        return StringUtils.fromString(dateStr);
    }

    private static boolean parseBoolean(String value) {
        if (value == null) {
            return false;
        }
        String lower = value.trim().toLowerCase();
        return "true".equals(lower) || "yes".equals(lower) || "1".equals(lower);
    }

    /**
     * Set a cell value from a Ballerina value.
     *
     * @param cell  The cell to set
     * @param value The Ballerina value
     */
    public static void setCellValue(Cell cell, Object value) {
        if (value == null) {
            cell.setBlank();
            return;
        }

        if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof io.ballerina.runtime.api.values.BDecimal) {
            cell.setCellValue(((io.ballerina.runtime.api.values.BDecimal) value).decimalValue().doubleValue());
        } else {
            // Default: convert to string
            cell.setCellValue(value.toString());
        }
    }
}
