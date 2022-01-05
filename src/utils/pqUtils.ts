// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import iconv from "iconv-lite";
import { URLS } from "../constants";
import {
    generateMashupXMLTemplate,
    generateCustomXmlFilePath,
} from "../generators";

const getBase64 = async (zip: JSZip): Promise<string | undefined> => {
    const { found, value } = await getCustumXmlFile(zip, URLS.DATA_MASHUP);
    if (!found) {
        throw new Error("DataMashup XML is not found");
    }
    return value;
};

const setBase64 = async (zip: JSZip, base64: string): Promise<void> => {
    const newXml = generateMashupXMLTemplate(base64);
    const encoded = iconv.encode(newXml, "UCS2", { addBOM: true });
    const { path } = await getCustumXmlFile(zip, URLS.DATA_MASHUP);
    zip.file(path, encoded);
};

const getCustumXmlFile = async (
    zip: JSZip,
    url: string,
    encoding = "UTF-16"
): Promise<{ found: boolean; path: string; value: string | undefined }> => {
    const parser: DOMParser = new DOMParser();
    let found = false;
    let path;
    let value;
    for (let i = 1; ; i++) {
        path = generateCustomXmlFilePath(i);
        const xmlValue = await zip.file(path)?.async("uint8array");

        if (xmlValue === undefined) {
            break;
        }

        const xmlString = iconv.decode(xmlValue.buffer as Buffer, encoding);
        const doc: Document = parser.parseFromString(xmlString, "text/xml");

        found = doc?.documentElement?.namespaceURI === url;

        if (found) {
            value = doc.documentElement.innerHTML;
            break;
        }
    }

    return { found, path, value };
};

export default {
    getBase64,
    setBase64,
    getCustumXmlFile,
};
