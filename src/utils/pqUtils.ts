// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { EmptyQueryNameErr, QueryNameMaxLengthErr, maxQueryLength, URLS, BOM, QueryNameInvalidCharsErr } from "./constants";
import { generateMashupXMLTemplate, generateCustomXmlFilePath } from "../generators";
import { Buffer } from "buffer";
import { DOMParser } from "xmldom-qsa";

type CustomXmlFile = {
    found: boolean;
    path: string;
    xmlString: string | undefined;
    value: string | null;
};

const getBase64 = async (zip: JSZip): Promise<string | null> => {
    const mashup = await getDataMashupFile(zip);
    return mashup.value;
};

const setBase64 = async (zip: JSZip, base64: string): Promise<void> => {
    const newXml = generateMashupXMLTemplate(base64);
    const encoded = Buffer.from(BOM + newXml, "ucs2");
    const mashup = await getDataMashupFile(zip);
    zip.file(mashup?.path, encoded);
};

const getDataMashupFile = async (zip: JSZip): Promise<CustomXmlFile> => {
    let mashup;

    for (const url of URLS.PQ) {
        const item = await getCustomXmlFile(zip, url);
        if (item.found) {
            mashup = item;
            break;
        }
    }

    if (!mashup) {
        throw new Error("DataMashup XML is not found");
    }

    return mashup;
};

const getCustomXmlFile = async (zip: JSZip, url: string): Promise<CustomXmlFile> => {
    const parser: DOMParser = new DOMParser();
    const itemsArray = await zip.file(/customXml\/item\d.xml/);

    if (!itemsArray || itemsArray.length === 0) {
        throw new Error("No customXml files were found!");
    }

    let found = false;
    let path: string;
    let xmlString: string | undefined;
    let value: string | null = null;

    for (let i = 1; i <= itemsArray.length; i++) {
        path = generateCustomXmlFilePath(i);
        const xmlValue = await zip.file(path)?.async("uint8array");

        if (xmlValue === undefined) {
            break;
        }

        const buffer: Buffer = Buffer.from(xmlValue);
        const encoding: string | null = detectEncoding(xmlValue);
        if (!encoding){
            throw new Error("Failed to detect xml encoding")
        }

        xmlString = buffer
            .toString(encoding as BufferEncoding)
            .replace(/^\ufeff/, "");
        const doc: Document = parser.parseFromString(xmlString, "text/xml");

        found = doc?.documentElement?.namespaceURI === url;

        if (found) {
            value = doc.documentElement.textContent;
            break;
        }
    }

    return { found, path: path!, xmlString: xmlString, value };
};

const queryNameHasInvalidChars = (queryName: string) => {
    const invalidQueryNameChars = ['"', "."];

    // Control characters as defined in Unicode
    for (let c = 0; c <= 0x001f; ++c) {
        invalidQueryNameChars.push(String.fromCharCode(c));
    }

    for (let c = 0x007f; c <= 0x009f; ++c) {
        invalidQueryNameChars.push(String.fromCharCode(c));
    }

    return queryName.split("").some((ch) => invalidQueryNameChars.indexOf(ch) !== -1);
};

const validateQueryName = (queryName: string): void => {
    if (queryName) {
        if (queryName.length > maxQueryLength) {
            throw new Error(QueryNameMaxLengthErr);
        }

        if (queryNameHasInvalidChars(queryName)) {
            throw new Error(QueryNameInvalidCharsErr);
        }
    }

    if (!queryName.trim()) {
        throw new Error(EmptyQueryNameErr);
    }
};

const detectEncoding = (xmlBytes: Uint8Array): string | null => {
    if (!xmlBytes){
        return null;
    }

    if (xmlBytes.length >= 3 && xmlBytes[0] === 0xEF && xmlBytes[1] === 0xBB && xmlBytes[2] === 0xBF) {
        return 'utf-8';
    }
    if (xmlBytes.length >= 3 && xmlBytes[0] === 0xFF && xmlBytes[1] === 0xFE) {
        return 'utf-16le';
    }

    if (xmlBytes.length >= 3 && xmlBytes[0] === 0xFE && xmlBytes[1] === 0xFF) {
        return 'utf-16be';
    }

    // Default to utf-8, that not required a BOM for encoding.
    return 'utf-8';
}

export default {
    getBase64,
    setBase64,
    getCustomXmlFile,
    getDataMashupFile,
    validateQueryName,
};
