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

package io.ballerina.lib.data.xlsx;

import io.ballerina.lib.data.xlsx.utils.DiagnosticLog;
import io.ballerina.lib.data.xlsx.xlsx.XlsxParser;
import io.ballerina.lib.data.xlsx.xlsx.XlsxWriter;
import io.ballerina.runtime.api.Environment;
import io.ballerina.runtime.api.values.BArray;
import io.ballerina.runtime.api.values.BMap;
import io.ballerina.runtime.api.values.BStream;
import io.ballerina.runtime.api.values.BString;
import io.ballerina.runtime.api.values.BTypedesc;

import java.io.ByteArrayOutputStream;
import java.io.IOException;

/**
 * Native entry point for Ballerina XLSX module.
 * This class provides the bridge between Ballerina and Java for Excel operations.
 *
 * @since 0.1.0
 */
public final class Native {

    private Native() {
        // Private constructor to prevent instantiation
    }

    /**
     * Parse XLSX bytes into a Ballerina array.
     *
     * @param xlsxBytes The XLSX file content as a byte array
     * @param options   Parsing options
     * @param typedesc  Target type descriptor
     * @return Parsed data as BArray or error
     */
    public static Object parseBytes(BArray xlsxBytes, BMap<BString, Object> options, BTypedesc typedesc) {
        byte[] bytes = xlsxBytes.getBytes();
        return XlsxParser.parseBytes(bytes, options, typedesc);
    }

    /**
     * Parse XLSX from a byte stream into a Ballerina array.
     * Note: Stream support requires async handling - for now, returns error.
     *
     * @param env        Ballerina environment
     * @param xlsxStream The XLSX content as a byte stream
     * @param options    Parsing options
     * @param typedesc   Target type descriptor
     * @return Parsed data as BArray or error
     */
    public static Object parseStream(Environment env, BStream xlsxStream,
                                     BMap<BString, Object> options, BTypedesc typedesc) {
        // TODO: Implement proper async stream handling
        return DiagnosticLog.error("parseStream is not yet implemented. Use parseBytes instead.");
    }

    /**
     * Convert Ballerina data to XLSX bytes.
     *
     * @param data    Data to write (record[] or string[][])
     * @param options Write options
     * @return XLSX bytes as byte[] or error
     */
    public static Object toBytes(BArray data, BMap<BString, Object> options) {
        return XlsxWriter.toBytes(data, options);
    }
}
