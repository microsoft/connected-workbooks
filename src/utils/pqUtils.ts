// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { EmptyQueryNameErr, QueryNameMaxLengthErr, maxQueryLength, URLS, BOM, QueryNameInvalidCharsErr, queryNameAlreadyExistsErr, defaults } from "./constants";
import { generateMashupXMLTemplate, generateCustomXmlFilePath } from "../generators";
import { Buffer } from "buffer";
import { ConnectionOnlyQueryInfo } from "../types";

type CustomXmlFile = {
    found: boolean;
    path: string;
    xmlString: string | undefined;
    value: string | undefined;
};

const getBase64 = async (zip: JSZip): Promise<string | undefined> => {
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

const getCustomXmlFile = async (zip: JSZip, url: string, encoding: BufferEncoding = "utf16le"): Promise<CustomXmlFile> => {
    const parser: DOMParser = new DOMParser();
    const itemsArray = await zip.file(/customXml\/item\d.xml/);

    if (!itemsArray || itemsArray.length === 0) {
        throw new Error("No customXml files were found!");
    }

    let found = false;
    let path: string;
    let xmlString: string | undefined;
    let value: string | undefined;

    for (let i = 1; i <= itemsArray.length; i++) {
        path = generateCustomXmlFilePath(i);
        const xmlValue = await zip.file(path)?.async("uint8array");

        if (xmlValue === undefined) {
            break;
        }

        xmlString = Buffer.from(xmlValue)
            .toString(encoding)
            .replace(/^\ufeff/, "");
        const doc: Document = parser.parseFromString(xmlString, "text/xml");

        found = doc?.documentElement?.namespaceURI === url;

        if (found) {
            value = doc.documentElement.innerHTML;
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

const validateMultipleQueryNames = (queries: ConnectionOnlyQueryInfo[],  loadedQueryName: string): string[] => {
    const queryNames: string[] = [];
    const cleanedLoadedQueryName: string = loadedQueryName.trim().toLowerCase();
    queries.forEach((query: ConnectionOnlyQueryInfo) => {
        if (query.queryName) {
            validateQueryName(query.queryName);
            const cleanedQueryName: string | undefined = query.queryName.trim().toLowerCase();
            if (queryNames.includes(cleanedQueryName) || cleanedQueryName === cleanedLoadedQueryName) {
                throw new Error(queryNameAlreadyExistsErr);
            }

            queryNames.push(cleanedQueryName);
        } 
    });
    
    return queryNames;
};

const assignQueryNames = (queries: ConnectionOnlyQueryInfo[], loadedQueryName: string, queryNames: string[]):  ConnectionOnlyQueryInfo[] => {
    // Generate unique name for queries without a name
    queries.forEach((query: ConnectionOnlyQueryInfo) => {
        if (!query.queryName) {
            query.queryName = generateUniqueQueryName(queryNames);
            queryNames.push(query.queryName);
        }
    });

    return queries;
};

const generateUniqueQueryName = (queryNames: string[]): string => {
    let queryName: string = defaults.queryName;
    let index: number = 2;
    while (queryNames.includes(queryName)) {
        queryName = defaults.queryNamePrefix + index++;
    }

    return queryName;
};

export default {
    getBase64,
    setBase64,
    getCustomXmlFile,
    getDataMashupFile,
    validateQueryName,
    assignQueryNames,
    validateMultipleQueryNames,
};
