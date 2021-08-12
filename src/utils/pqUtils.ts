// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import iconv from "iconv-lite";
import { pqCustomXmlPath } from "../constants";
import { generateMashupXMLTemplate } from "../generators";

const getBase64 = async (zip: JSZip): Promise<string> => {
    const xmlValue = await zip.file(pqCustomXmlPath)?.async("uint8array");
    if (xmlValue === undefined) {
        throw new Error("PQ document wasn't found in zip");
    }
    const xmlString = iconv.decode(xmlValue.buffer as Buffer, "UTF-16");
    const parser: DOMParser = new DOMParser();
    const doc: Document = parser.parseFromString(xmlString, "text/xml");
    const result = doc.childNodes[0].textContent;
    if (result === null) {
        throw Error("Base64 wasn't found in zip");
    }
    return result;
};

const setBase64 = (zip: JSZip, base64: string): void => {
    const newXml = generateMashupXMLTemplate(base64);
    const encoded = iconv.encode(newXml, "UCS2", { addBOM: true });
    zip.file(pqCustomXmlPath, encoded);
};

export default {
    getBase64,
    setBase64,
};
