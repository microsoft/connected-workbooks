// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, test, expect } from '@jest/globals';
import { arrayUtils } from "../src/utils/";


test("getInt32Buffer test", () => {
    const size = 4;
    const val = 4;

    const packageSizeBuffer = new ArrayBuffer(size);
    new DataView(packageSizeBuffer).setInt32(0, val, true);
    const expected = new Uint8Array(packageSizeBuffer);

    const actual = arrayUtils.getInt32Buffer(size);

    expect(actual).toStrictEqual(expected);
});

test("concatArrays test", () => {
    const uIntArr1 = new Uint8Array(4).fill(5);
    const uIntArr2 = new Uint8Array(2).fill(10);
    const expected = new Uint8Array(6);

    expected.set(uIntArr1, 0);
    expected.set(uIntArr2, 4);

    const actual = arrayUtils.concatArrays(uIntArr1, uIntArr2);

    expect(actual).toStrictEqual(expected);
});

describe("base64ToUint8Array tests", () => {
    test("basic base64 decoding", () => {
        const base64 = "SGVsbG8gV29ybGQ="; // "Hello World"
        const [result, dataView] = arrayUtils.base64ToUint8Array(base64);
        const expected = new Uint8Array([72, 101, 108, 108, 111, 32, 87, 111, 114, 108, 100]);
        
        expect(result).toStrictEqual(expected);
        expect(dataView).toBeInstanceOf(DataView);
        expect(dataView.byteLength).toBe(expected.length);
    });

    test("base64 with single padding", () => {
        const base64 = "SGVsbG8="; // "Hello"
        const [result] = arrayUtils.base64ToUint8Array(base64);
        const expected = new Uint8Array([72, 101, 108, 108, 111]);
        
        expect(result).toStrictEqual(expected);
    });

    test("base64 with double padding", () => {
        const base64 = "QQ=="; // Single character "A"
        const [result] = arrayUtils.base64ToUint8Array(base64);
        const expected = new Uint8Array([65]); // "A" = 65
        
        expect(result).toStrictEqual(expected);
    });

    test("base64 with no padding", () => {
        const base64 = "SGVsbG8h"; // "Hello!"
        const [result] = arrayUtils.base64ToUint8Array(base64);
        const expected = new Uint8Array([72, 101, 108, 108, 111, 33]);
        
        expect(result).toStrictEqual(expected);
    });

    test("empty base64 string", () => {
        const base64 = "";
        const [result] = arrayUtils.base64ToUint8Array(base64);
        
        expect(result).toStrictEqual(new Uint8Array(0));
    });

    test("base64 with whitespace", () => {
        const base64 = "SGVs bG8g V29y bGQ="; // "Hello World" with spaces
        const [result] = arrayUtils.base64ToUint8Array(base64);
        const expected = new Uint8Array([72, 101, 108, 108, 111, 32, 87, 111, 114, 108, 100]);
        
        expect(result).toStrictEqual(expected);
    });
});

describe("uint8ArrayToBase64 tests", () => {
    test("basic uint8array to base64", () => {
        const input = new Uint8Array([72, 101, 108, 108, 111, 32, 87, 111, 114, 108, 100]); // "Hello World"
        const result = arrayUtils.uint8ArrayToBase64(input);
        
        expect(result).toBe("SGVsbG8gV29ybGQ=");
    });

    test("uint8array requiring single padding", () => {
        const input = new Uint8Array([72, 101, 108, 108, 111]); // "Hello"
        const result = arrayUtils.uint8ArrayToBase64(input);
        
        expect(result).toBe("SGVsbG8=");
    });

    test("uint8array requiring double padding", () => {
        const input = new Uint8Array([65]); // Single byte "A"
        const result = arrayUtils.uint8ArrayToBase64(input);
        
        expect(result).toBe("QQ==");
    });

    test("uint8array with two bytes (single padding)", () => {
        const input = new Uint8Array([72, 105]); // "Hi"
        const result = arrayUtils.uint8ArrayToBase64(input);
        
        expect(result).toBe("SGk=");
    });

    test("uint8array with no padding needed", () => {
        const input = new Uint8Array([72, 101, 108, 108, 111, 33]); // "Hello!"
        const result = arrayUtils.uint8ArrayToBase64(input);
        
        expect(result).toBe("SGVsbG8h");
    });

    test("empty uint8array", () => {
        const input = new Uint8Array(0);
        const result = arrayUtils.uint8ArrayToBase64(input);
        
        expect(result).toBe("");
    });

    test("basic ASCII string", () => {
        const input = "Hello";
        const result = arrayUtils.encodeStringToUCS2(input);
        // H=72, e=101, l=108, l=108, o=111 in little-endian UCS-2
        const expected = new Uint8Array([72, 0, 101, 0, 108, 0, 108, 0, 111, 0]);
        
        expect(result).toStrictEqual(expected);
    });

    test("empty string", () => {
        const input = "";
        const result = arrayUtils.encodeStringToUCS2(input);
        
        expect(result).toStrictEqual(new Uint8Array(0));
    });

    test("string with unicode characters", () => {
        const input = "A€"; // A=65, Euro=8364
        const result = arrayUtils.encodeStringToUCS2(input);
        // A=65 (0x41), €=8364 (0x20AC) in little-endian
        const expected = new Uint8Array([65, 0, 172, 32]);
        
        expect(result).toStrictEqual(expected);
    });

    test("string with high unicode character", () => {
        const input = "🙂"; // Emoji - this will be encoded as surrogate pair
        const result = arrayUtils.encodeStringToUCS2(input);
        
        expect(result).toHaveLength(4); // 2 characters * 2 bytes each (surrogate pair)
    });

    test("single character", () => {
        const input = "X";
        const result = arrayUtils.encodeStringToUCS2(input);
        const expected = new Uint8Array([88, 0]); // X=88
        
        expect(result).toStrictEqual(expected);
    });
});

describe("base64 and uint8Array roundtrip tests", () => {
    test("empty data roundtrip", () => {
        const original = new Uint8Array(0);
        const base64 = arrayUtils.uint8ArrayToBase64(original);
        const [decoded] = arrayUtils.base64ToUint8Array(base64);
        
        expect(decoded).toStrictEqual(original);
        expect(base64).toBe("");
    });

    test("single byte roundtrip", () => {
        const original = new Uint8Array([42]);
        const base64 = arrayUtils.uint8ArrayToBase64(original);
        const [decoded] = arrayUtils.base64ToUint8Array(base64);
        
        expect(decoded).toStrictEqual(original);
        expect(base64.endsWith("==")).toBe(true); // Should have double padding
    });

    test("two bytes roundtrip", () => {
        const original = new Uint8Array([42, 123]);
        const base64 = arrayUtils.uint8ArrayToBase64(original);
        const [decoded] = arrayUtils.base64ToUint8Array(base64);
        
        expect(decoded).toStrictEqual(original);
        expect(base64.endsWith("=")).toBe(true); // Should have single padding
        expect(base64.endsWith("==")).toBe(false);
    });

    test("three bytes roundtrip (no padding)", () => {
        const original = new Uint8Array([42, 123, 200]);
        const base64 = arrayUtils.uint8ArrayToBase64(original);
        const [decoded] = arrayUtils.base64ToUint8Array(base64);
        
        expect(decoded).toStrictEqual(original);
        expect(base64.includes("=")).toBe(false); // No padding needed
    });

    test("random data roundtrip", () => {
        for (let size = 0; size < 100; size++) {
            const original = new Uint8Array(size);
            for (let i = 0; i < size; i++) {
                original[i] = Math.floor(Math.random() * 256);
            }
            
            const base64 = arrayUtils.uint8ArrayToBase64(original);
            const [decoded, dataView] = arrayUtils.base64ToUint8Array(base64);
            
            // Verify roundtrip
            expect(decoded).toStrictEqual(original);
            // Verify DataView is properly constructed
            expect(dataView.byteLength).toBe(original.length);
        }
    });

    test("base64 with whitespace handling", () => {
        const original = new Uint8Array([1, 2, 3, 4, 5]);
        const base64Clean = arrayUtils.uint8ArrayToBase64(original);
        const base64WithWhitespace = base64Clean.split('').join(' '); // Add spaces
        
        const [decoded] = arrayUtils.base64ToUint8Array(base64WithWhitespace);
        expect(decoded).toStrictEqual(original);
    });

    test("base64 properties", () => {
        const testCases = [
            new Uint8Array([]),
            new Uint8Array([1]),
            new Uint8Array([1, 2]),
            new Uint8Array([1, 2, 3]),
            new Uint8Array([1, 2, 3, 4]),
            new Uint8Array([255, 255, 255])
        ];

        testCases.forEach(original => {
            const base64 = arrayUtils.uint8ArrayToBase64(original);
            
            // Base64 should only contain valid characters
            expect(base64).toMatch(/^[A-Za-z0-9+/]*={0,2}$/);
            
            // Length should be correct (padded to multiple of 4)
            expect(base64.length % 4).toBe(0);
            
            // Roundtrip should work
            const [decoded] = arrayUtils.base64ToUint8Array(base64);
            expect(decoded).toStrictEqual(original);
        });
    });
});

describe("encodeStringToUCS2 tests", () => {
    test("empty string", () => {
        const result = arrayUtils.encodeStringToUCS2("");
        expect(result).toStrictEqual(new Uint8Array(0));
    });

    test("string length matches byte length", () => {
        const testStrings = ["A", "AB", "Hello", "Hello World"];
        
        testStrings.forEach(str => {
            const result = arrayUtils.encodeStringToUCS2(str);
            expect(result.length).toBe(str.length * 2);
        });
    });

    test("ASCII characters are encoded correctly", () => {
        // Test that we can decode what we encode using built-in TextDecoder
        const testStrings = ["A", "Hello", "Test123"];
        
        testStrings.forEach(str => {
            const encoded = arrayUtils.encodeStringToUCS2(str);
            const decoded = new TextDecoder('utf-16le').decode(encoded);
            expect(decoded).toBe(str);
        });
    });

    test("unicode characters are handled", () => {
        const testStrings = ["€", "🙂", "café"];
        
        testStrings.forEach(str => {
            const encoded = arrayUtils.encodeStringToUCS2(str);
            const decoded = new TextDecoder('utf-16le').decode(encoded);
            expect(decoded).toBe(str);
        });
    });

    test("little-endian byte order", () => {
        const result = arrayUtils.encodeStringToUCS2("A"); // A = 65 = 0x41
        expect(result).toStrictEqual(new Uint8Array([0x41, 0x00])); // Little-endian
    });
});

describe("decodeXml tests", () => {
    const sampleXml = '<?xml version="1.0"?><root>test</root>';

    test("UTF-8 without BOM", () => {
        const xmlBytes = new TextEncoder().encode(sampleXml);
        const result = arrayUtils.decodeXml(xmlBytes);
        expect(result).toBe(sampleXml);
    });

    test("UTF-8 with BOM is stripped", () => {
        const xmlBytes = new Uint8Array([0xEF, 0xBB, 0xBF, ...new TextEncoder().encode(sampleXml)]);
        const result = arrayUtils.decodeXml(xmlBytes);
        expect(result).toBe(sampleXml);
        expect(result.charCodeAt(0)).not.toBe(0xFEFF); // BOM should be removed
    });

    test("UTF-16LE with BOM", () => {
        // Create UTF-16LE with BOM manually
        const utf16Bytes = new Uint8Array(2 + sampleXml.length * 2);
        utf16Bytes[0] = 0xFF; // BOM
        utf16Bytes[1] = 0xFE;
        
        // Encode each character as little-endian UTF-16
        for (let i = 0; i < sampleXml.length; i++) {
            const code = sampleXml.charCodeAt(i);
            utf16Bytes[2 + i * 2] = code & 0xff;
            utf16Bytes[2 + i * 2 + 1] = code >> 8;
        }
        
        const result = arrayUtils.decodeXml(utf16Bytes);
        expect(result).toBe(sampleXml);
    });

    test("UTF-16BE with BOM", () => {
        // Create UTF-16BE with BOM manually  
        const utf16Bytes = new Uint8Array(2 + sampleXml.length * 2);
        utf16Bytes[0] = 0xFE; // BOM
        utf16Bytes[1] = 0xFF;
        
        // Encode each character as big-endian UTF-16
        for (let i = 0; i < sampleXml.length; i++) {
            const code = sampleXml.charCodeAt(i);
            utf16Bytes[2 + i * 2] = code >> 8;
            utf16Bytes[2 + i * 2 + 1] = code & 0xff;
        }
        
        const result = arrayUtils.decodeXml(utf16Bytes);
        expect(result).toBe(sampleXml);
    });

    test("empty bytes throws error", () => {
        expect(() => arrayUtils.decodeXml(new Uint8Array(0))).toThrow("Failed to detect xml encoding");
    });

    test("invalid UTF-16BE throws error", () => {
        const invalidBytes = new Uint8Array([0xFE, 0xFF, 0x00]); // BOM + odd length
        expect(() => arrayUtils.decodeXml(invalidBytes)).toThrow("Invalid UTF-16BE byte array");
    });

    test("encoding detection works correctly", () => {
        // Test that different BOMs are detected and handled
        const testCases = [
            { bytes: new Uint8Array([0xEF, 0xBB, 0xBF, 65]), name: "UTF-8 BOM" },
            { bytes: new Uint8Array([0xFF, 0xFE, 65, 0]), name: "UTF-16LE BOM" },
            { bytes: new Uint8Array([0xFE, 0xFF, 0, 65]), name: "UTF-16BE BOM" },
            { bytes: new Uint8Array([65]), name: "No BOM (defaults to UTF-8)" }
        ];

        testCases.forEach(({ bytes }) => {
            expect(() => arrayUtils.decodeXml(bytes)).not.toThrow();
        });
    });
});