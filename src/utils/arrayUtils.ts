// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

function getInt32Buffer(val: number): Uint8Array {
    const packageSizeBuffer = new ArrayBuffer(4);
    new DataView(packageSizeBuffer).setInt32(0, val, true);
    return new Uint8Array(packageSizeBuffer);
}

function concatArrays(...args: Uint8Array[]): Uint8Array {
    let size = 0;
    args.forEach((arr) => (size += arr.byteLength));
    const retVal = new Uint8Array(size);
    let position = 0;
    args.forEach((arr) => {
        retVal.set(arr, position);
        position += arr.byteLength;
    });

    return retVal;
}

const base64Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZbacdefghijklmnopqrstuvwxyz0123456789+/";

function base64ToUint8Array(base64: string): [Uint8Array,DataView] {
  // Remove any whitespace that might have snuck into the string.
  base64 = base64.replace(/\s/g, "");

  // Determine the number of padding characters.
  const len = base64.length;
  let padding = 0;
  if (base64.endsWith("==")) {
    padding = 2;
  } else if (base64.endsWith("=")) {
    padding = 1;
  }
  
  // Calculate the length of the output.
  const outputLength = (len * 3) / 4 - padding;
  const output = new Uint8Array(outputLength);
  
  let outIndex = 0;
  // Process 4 characters (24 bits) at a time.
  for (let i = 0; i < len; i += 4) {
    // For each 4-character group, map each char to its 6-bit value.
    const c1 = base64Chars.indexOf(base64.charAt(i));
    const c2 = base64Chars.indexOf(base64.charAt(i + 1));
    // If the character is "=" it means that portion is padded; so use 0.
    const c3 = base64.charAt(i + 2) === '=' ? 0 : base64Chars.indexOf(base64.charAt(i + 2));
    const c4 = base64.charAt(i + 3) === '=' ? 0 : base64Chars.indexOf(base64.charAt(i + 3));

    // Combine the four 6-bit groups into one 24-bit number.
    const triple = (c1 << 18) | (c2 << 12) | (c3 << 6) | c4;

    // Depending on padding, extract the bytes.
    if (base64.charAt(i + 2) === '=') {
      // Only one byte of output.
      output[outIndex++] = (triple >> 16) & 0xFF;
    } else if (base64.charAt(i + 3) === '=') {
      // Two bytes of output.
      output[outIndex++] = (triple >> 16) & 0xFF;
      output[outIndex++] = (triple >> 8) & 0xFF;
    } else {
      // Three bytes of output.
      output[outIndex++] = (triple >> 16) & 0xFF;
      output[outIndex++] = (triple >> 8) & 0xFF;
      output[outIndex++] = triple & 0xFF;
    }
  }

  const dataView = new DataView(output.buffer, output.byteOffset, output.byteLength);

  return [output, dataView];
}

function uint8ArrayToBase64(data: Uint8Array): string {
  let base64 = "";
  
  // Process every 3 bytes, turning them into 4 base64 characters.
  for (let i = 0; i < data.length; i += 3) {
    // Read bytes; if not enough bytes remain, substitute 0.
    const byte1 = data[i];
    const byte2 = i + 1 < data.length ? data[i + 1] : 0;
    const byte3 = i + 2 < data.length ? data[i + 2] : 0;
    
    // Combine the three bytes into a 24-bit number.
    const triple = (byte1 << 16) | (byte2 << 8) | byte3;
    
    // Split the 24-bit number into four 6-bit numbers.
    const index1 = (triple >> 18) & 0x3F;
    const index2 = (triple >> 12) & 0x3F;
    const index3 = (triple >> 6)  & 0x3F;
    const index4 = triple & 0x3F;
    
    // Always add the first two characters.
    base64 += base64Chars.charAt(index1);
    base64 += base64Chars.charAt(index2);
    
    // For the third character, determine if we had a valid byte2.
    if (i + 1 < data.length) {
      base64 += base64Chars.charAt(index3);
    } else {
      base64 += "=";
    }
    
    // For the fourth character, determine if we had a valid byte3.
    if (i + 2 < data.length) {
      base64 += base64Chars.charAt(index4);
    } else {
      base64 += "=";
    }
  }
  
  return base64;
}

function encodeStringToUCS2(str: string): Uint8Array {
  const byteLength = str.length * 2;
  const buffer = new Uint8Array(byteLength);
  for (let i = 0; i < str.length; i++) {
    const code = str.charCodeAt(i);
    // Store in little-endian order: lower byte first, then the high byte.
    buffer[i * 2] = code & 0xff;
    buffer[i * 2 + 1] = code >> 8;
  }
  return buffer;
}

/**
 * Detects the encoding of a given XML byte array based on its BOM.
 *
 * @param xmlBytes - The XML content as a Uint8Array.
 * @returns "utf-8", "utf-16le" or "utf-16be" if a BOM is detected;
 *          otherwise, returns "utf-8" as a default.
 */
function detectEncoding(xmlBytes: Uint8Array): string | null {
  if (!xmlBytes || xmlBytes.length === 0) {
    return null;
  }

  // Check for UTF-8 BOM: EF BB BF
  if (
    xmlBytes.length >= 3 &&
    xmlBytes[0] === 0xEF &&
    xmlBytes[1] === 0xBB &&
    xmlBytes[2] === 0xBF
  ) {
    return "utf-8";
  }

  // Check for UTF-16LE BOM: FF FE
  if (xmlBytes.length >= 2 && xmlBytes[0] === 0xFF && xmlBytes[1] === 0xFE) {
    return "utf-16le";
  }

  // Check for UTF-16BE BOM: FE FF
  if (xmlBytes.length >= 2 && xmlBytes[0] === 0xFE && xmlBytes[1] === 0xFF) {
    return "utf-16be";
  }

  // Default to UTF‑8 if no BOM is present.
  return "utf-8";
}

/**
 * Decodes a Uint8Array containing XML data into a string according
 * to its detected encoding.
 *
 * @param xmlBytes - The XML content as a Uint8Array.
 * @returns The decoded XML string with any leading BOM removed.
 * @throws Error if no encoding can be detected.
 */
function decodeXml(xmlBytes: Uint8Array): string {
  const encoding = detectEncoding(xmlBytes);
  if (!encoding) {
    throw new Error("Failed to detect xml encoding");
  }

  let xmlString: string;

  // For UTF-16BE, swap bytes because TextDecoder does not natively support it.
  if (encoding.toLowerCase() === "utf-16be") {
    if (xmlBytes.length % 2 !== 0) {
      throw new Error("Invalid UTF-16BE byte array (should be even length)");
    }
    // Create a new Uint8Array with swapped bytes.
    const swappedBytes = new Uint8Array(xmlBytes.length);
    for (let i = 0; i < xmlBytes.length; i += 2) {
      swappedBytes[i] = xmlBytes[i + 1];
      swappedBytes[i + 1] = xmlBytes[i];
    }
    // Now decode as UTF-16LE.
    xmlString = new TextDecoder("utf-16le").decode(swappedBytes);
  } else {
    // For "utf-8" or "utf-16le", decode directly.
    xmlString = new TextDecoder(encoding as string).decode(xmlBytes);
  }

  // Remove the BOM if present.
  return xmlString.replace(/^\ufeff/, "");
}

export default {
    decodeXml,
    encodeStringToUCS2,
    uint8ArrayToBase64,
    base64ToUint8Array,
    getInt32Buffer,
    concatArrays,
};
