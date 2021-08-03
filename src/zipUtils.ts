import JSZipModule from "jszip";
import WorkbookTemplate from "./workbookTemplate";
import iconv from "iconv-lite";
import { pqCustomXmlPath, section1mPath } from "./constants";
import {
    generateMashupXMLTemplate,
    generateSection1mString,
} from "./generators";

export type JSZip = JSZipModule;

const loadAsync = JSZipModule.loadAsync;

const loadAsyncDefaultTemplate = async (): Promise<JSZip> =>
    await JSZipModule.loadAsync(
        WorkbookTemplate.SIMPLE_QUERY_WORKBOOK_TEMPLATE,
        {
            base64: true,
        }
    );

const loadAsyncTemplate = async (template: string): Promise<JSZip> =>
    await JSZipModule.loadAsync(template, {
        base64: true,
    });

const chackAndgetSection1m = async (zip: JSZip): Promise<string> => {
    const section1m = zip.file("Formulas/Section1.m")?.async("text");
    if (!section1m) {
        throw new Error("Formula section wasn't found in template");
    }

    return section1m;
};

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

const setSection1m = (queryName: string, query: string, zip: JSZip): void => {
    const newSection1m = generateSection1mString(queryName, query);

    zip.file(section1mPath, newSection1m, {
        compression: "",
    });
};

export default {
    loadAsyncDefaultTemplate,
    chackAndgetSection1m,
    loadAsyncTemplate,
    setSection1m,
    loadAsync,
    getBase64,
    setBase64,
};
