/* eslint-disable @typescript-eslint/no-var-requires */
const fs = require("fs/promises");
const path = require("path");
const JSZip = require("jszip");
const { DOMParser } = require("@xmldom/xmldom");

const DATA_MASHUP_REGEX = /<DataMashup[^>]*>([\s\S]+?)<\/DataMashup>/i;

function decodeXmlBuffer(buffer) {
    if (buffer.length >= 2 && buffer[0] === 0xff && buffer[1] === 0xfe) {
        return buffer.toString("utf16le").replace(/^\ufeff/, "");
    }

    return buffer.toString("utf8").replace(/^\ufeff/, "");
}

async function loadWorkbookZip(workbookPath) {
    const absolutePath = path.resolve(workbookPath);
    const buffer = await fs.readFile(absolutePath);
    return JSZip.loadAsync(buffer);
}

function extractQueryName(sectionContent) {
    const normalized = sectionContent.replace(/\r/g, "");
    const match = normalized.match(/shared\s+(?:#")?([^"\n]+)"?\s*=/i);
    return match ? match[1].trim() : "<unknown>";
}

async function readMashup(zip) {
    const entry = zip.file("customXml/item1.xml");
    if (!entry) {
        throw new Error("customXml/item1.xml not found");
    }

    const xmlBuffer = await entry.async("nodebuffer");
    const mashupXml = decodeXmlBuffer(xmlBuffer);
    const match = mashupXml.match(DATA_MASHUP_REGEX);
    if (!match) {
        throw new Error("DataMashup payload missing");
    }

    const payloadBuffer = Buffer.from(match[1].trim(), "base64");
    let offset = 0;
    const readInt32 = () => {
        const value = payloadBuffer.readInt32LE(offset);
        offset += 4;
        return value;
    };

    offset += 4; // version
    const packageSize = readInt32();
    const packageBytes = payloadBuffer.subarray(offset, offset + packageSize);
    offset += packageSize;
    const permissionsSize = readInt32();
    offset += permissionsSize;
    const metadataSize = readInt32();
    const metadataBytes = payloadBuffer.subarray(offset, offset + metadataSize);

    const packageZip = await JSZip.loadAsync(packageBytes);
    const sectionEntry = packageZip.file("Formulas/Section1.m");
    if (!sectionEntry) {
        throw new Error("Formulas/Section1.m missing");
    }

    const sectionContent = (await sectionEntry.async("text")).trim();
    const queryName = extractQueryName(sectionContent);

    let metadataOffset = 4; // metadata version
    const metadataXmlLength = metadataBytes.readInt32LE(metadataOffset);
    metadataOffset += 4;
    const metadataXml = metadataBytes.subarray(metadataOffset, metadataOffset + metadataXmlLength).toString("utf8");

    return {
        queryName,
        sectionContent,
        metadataBytesLength: metadataBytes.length,
        metadataXml,
    };
}

function parseSharedStrings(xml) {
    if (!xml) {
        return [];
    }

    const doc = new DOMParser().parseFromString(xml, "text/xml");
    const nodes = doc.getElementsByTagName("t");
    const values = [];
    for (let i = 0; i < nodes.length; i++) {
        values.push((nodes[i].textContent || "").trim());
    }

    return values;
}

function parseMetadataPaths(metadataXml) {
    const doc = new DOMParser().parseFromString(metadataXml, "text/xml");
    const nodes = doc.getElementsByTagName("ItemPath");
    const paths = [];
    for (let i = 0; i < nodes.length; i++) {
        paths.push(nodes[i].textContent || "");
    }

    return paths;
}

async function extractWorkbookDetails(workbookPath) {
    const zip = await loadWorkbookZip(workbookPath);
    const mashup = await readMashup(zip);

    const connectionsEntry = zip.file("xl/connections.xml");
    if (!connectionsEntry) {
        throw new Error("xl/connections.xml not found");
    }

    const connectionsXml = await connectionsEntry.async("text");
    const connectionDoc = new DOMParser().parseFromString(connectionsXml, "text/xml");
    const connectionNode = connectionDoc.getElementsByTagName("connection")[0];
    const dbPrNode = connectionDoc.getElementsByTagName("dbPr")[0];

    const connection = {
        id: connectionNode?.getAttribute("id") || "",
        name: connectionNode?.getAttribute("name") || "",
        description: connectionNode?.getAttribute("description") || "",
        refreshOnLoad: dbPrNode?.getAttribute("refreshOnLoad") || "",
        location: dbPrNode?.getAttribute("connection") || "",
        command: dbPrNode?.getAttribute("command") || "",
    };

    const sharedStringsEntry = zip.file("xl/sharedStrings.xml");
    const sharedStringsXml = sharedStringsEntry ? await sharedStringsEntry.async("text") : "";

    return {
        workbookPath: path.resolve(workbookPath),
        queryName: mashup.queryName,
        metadataXml: mashup.metadataXml,
        metadataBytesLength: mashup.metadataBytesLength,
        sectionContent: mashup.sectionContent,
        metadataItemPaths: parseMetadataPaths(mashup.metadataXml),
        connection,
        sharedStrings: parseSharedStrings(sharedStringsXml),
    };
}

async function extractMashupInfo(workbookPath) {
    const details = await extractWorkbookDetails(workbookPath);
    return {
        workbookPath: details.workbookPath,
        queryName: details.queryName,
        sectionContent: details.sectionContent,
        metadataBytesLength: details.metadataBytesLength,
        metadataXml: details.metadataXml,
    };
}

module.exports = {
    extractMashupInfo,
    extractWorkbookDetails,
};
