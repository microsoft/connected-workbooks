// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import iconv from "iconv-lite";
import { URLS } from "../constants";
import { generateMashupXMLTemplate, generateCustomXmlFilePath } from "../generators";

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
    const encoded = iconv.encode(newXml, "UCS2", { addBOM: true });
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

const getCustomXmlFile = async (zip: JSZip, url: string, encoding = "UTF-16"): Promise<CustomXmlFile> => {
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

        xmlString = iconv.decode(xmlValue.buffer as Buffer, encoding);
        const doc: Document = parser.parseFromString(xmlString, "text/xml");

        found = doc?.documentElement?.namespaceURI === url;

        if (found) {
            value = doc.documentElement.innerHTML;
            break;
        }
    }

    return { found, path: path!, xmlString: xmlString, value };
};

export default {
    getBase64,
    setBase64,
    getCustomXmlFile,
    getDataMashupFile,
};
