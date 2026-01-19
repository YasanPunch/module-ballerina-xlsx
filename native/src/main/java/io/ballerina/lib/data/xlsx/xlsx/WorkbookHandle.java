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

import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.runtime.api.creators.ValueCreator;
import io.ballerina.runtime.api.utils.StringUtils;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BObject;
import io.ballerina.runtime.api.values.BString;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * Native handle for Apache POI Workbook.
 * This class wraps the POI Workbook and provides methods called from Ballerina.
 *
 * @since 0.1.0
 */
public final class WorkbookHandle {

    private static final String WORKBOOK_NATIVE_KEY = "workbookNative";

    private WorkbookHandle() {
        // Private constructor to prevent instantiation
    }

    /**
     * Open a workbook from bytes.
     *
     * @param workbookObj Ballerina Workbook object
     * @param data        XLSX file bytes
     * @return null on success, error on failure
     */
    public static Object openWorkbook(BObject workbookObj, BArray data) {
        try {
            ByteArrayInputStream bis = new ByteArrayInputStream(data.getBytes());
            Workbook workbook = WorkbookFactory.create(bis);
            workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, workbook);
            return null;
        } catch (IOException e) {
            return DiagnosticLog.parseError("Failed to open workbook: " + e.getMessage());
        } catch (Exception e) {
            return DiagnosticLog.error("Error opening workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Create a new empty workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return null on success, error on failure
     */
    public static Object createNewWorkbook(BObject workbookObj) {
        try {
            Workbook workbook = new XSSFWorkbook();
            workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, workbook);
            return null;
        } catch (Exception e) {
            return DiagnosticLog.error("Error creating workbook: " + e.getMessage(), e);
        }
    }

    /**
     * Get all sheet names from workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Array of sheet names
     */
    public static BArray getSheetNames(BObject workbookObj) {
        Workbook workbook = getWorkbook(workbookObj);
        int sheetCount = workbook.getNumberOfSheets();

        BString[] names = new BString[sheetCount];
        for (int i = 0; i < sheetCount; i++) {
            names[i] = StringUtils.fromString(workbook.getSheetName(i));
        }

        return ValueCreator.createArrayValue(names);
    }

    /**
     * Get number of sheets in workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @return Sheet count
     */
    public static long getSheetCount(BObject workbookObj) {
        Workbook workbook = getWorkbook(workbookObj);
        return workbook.getNumberOfSheets();
    }

    /**
     * Get a sheet by name.
     *
     * @param workbookObj Ballerina Workbook object
     * @param sheetObj    Ballerina Sheet object to initialize
     * @param name        Sheet name
     * @return null on success, error if not found
     */
    public static Object getSheet(BObject workbookObj, BObject sheetObj, BString name) {
        Workbook workbook = getWorkbook(workbookObj);
        String sheetName = name.getValue();

        Sheet sheet = workbook.getSheet(sheetName);
        if (sheet == null) {
            return DiagnosticLog.sheetNotFoundError(sheetName);
        }

        SheetHandle.initSheet(sheetObj, sheet);
        return null;
    }

    /**
     * Get a sheet by index.
     *
     * @param workbookObj Ballerina Workbook object
     * @param sheetObj    Ballerina Sheet object to initialize
     * @param index       Sheet index (0-based)
     * @return null on success, error if index out of range
     */
    public static Object getSheetByIndex(BObject workbookObj, BObject sheetObj, long index) {
        Workbook workbook = getWorkbook(workbookObj);
        int idx = (int) index;

        if (idx < 0 || idx >= workbook.getNumberOfSheets()) {
            return DiagnosticLog.error("Sheet index " + idx + " out of range (0-" +
                    (workbook.getNumberOfSheets() - 1) + ")");
        }

        Sheet sheet = workbook.getSheetAt(idx);
        SheetHandle.initSheet(sheetObj, sheet);
        return null;
    }

    /**
     * Create a new sheet in the workbook.
     *
     * @param workbookObj Ballerina Workbook object
     * @param sheetObj    Ballerina Sheet object to initialize
     * @param name        Name for the new sheet
     * @return null on success, error on failure
     */
    public static Object createSheet(BObject workbookObj, BObject sheetObj, BString name) {
        Workbook workbook = getWorkbook(workbookObj);
        String sheetName = name.getValue();

        try {
            // Check if sheet already exists
            if (workbook.getSheet(sheetName) != null) {
                return DiagnosticLog.error("Sheet '" + sheetName + "' already exists");
            }

            Sheet sheet = workbook.createSheet(sheetName);
            SheetHandle.initSheet(sheetObj, sheet);
            return null;

        } catch (Exception e) {
            return DiagnosticLog.error("Error creating sheet: " + e.getMessage(), e);
        }
    }

    /**
     * Convert workbook to bytes.
     *
     * @param workbookObj Ballerina Workbook object
     * @return XLSX bytes, or error
     */
    public static Object toBytes(BObject workbookObj) {
        Workbook workbook = getWorkbook(workbookObj);

        try {
            ByteArrayOutputStream bos = new ByteArrayOutputStream();
            workbook.write(bos);
            return ValueCreator.createArrayValue(bos.toByteArray());
        } catch (IOException e) {
            return DiagnosticLog.error("Failed to convert workbook to bytes: " + e.getMessage(), e);
        }
    }

    /**
     * Close the workbook and release resources.
     *
     * @param workbookObj Ballerina Workbook object
     * @return null on success, error on failure
     */
    public static Object close(BObject workbookObj) {
        Workbook workbook = (Workbook) workbookObj.getNativeData(WORKBOOK_NATIVE_KEY);
        if (workbook != null) {
            try {
                workbook.close();
                workbookObj.addNativeData(WORKBOOK_NATIVE_KEY, null);
            } catch (IOException e) {
                return DiagnosticLog.error("Error closing workbook: " + e.getMessage(), e);
            }
        }
        return null;
    }

    /**
     * Get the native Workbook from Ballerina object.
     */
    static Workbook getWorkbook(BObject workbookObj) {
        return (Workbook) workbookObj.getNativeData(WORKBOOK_NATIVE_KEY);
    }
}
