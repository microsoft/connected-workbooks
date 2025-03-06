// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import {
    section1mPath,
    defaults,
    uint8ArrayType,
    emptyValue,
    textResultType,
    xmlTextResultType,
    element,
    section1PathPrefix,
    divider,
    elementAttributes,
    elementAttributesValues,
    Errors,
} from "./constants";
import { arrayUtils } from ".";
import { Metadata } from "../types";
import { DOMParser, XMLSerializer } from "xmldom-qsa";
import { Buffer } from "buffer";

export const replaceSingleQuery = async (base64Str: string, queryName: string, queryMashupDoc: string): Promise<string> => {
    const { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } = getPackageComponents(base64Str);
    const newPackageBuffer: Uint8Array = await editSingleQueryPackage(packageOPC, queryMashupDoc);
    const packageSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(newPackageBuffer.byteLength);
    const permissionsSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(permissionsSize);
    const newMetadataBuffer: Uint8Array = editSingleQueryMetadata(metadata, { queryName });
    const metadataSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
    const newMashup: Uint8Array = arrayUtils.concatArrays(
        version,
        packageSizeBuffer,
        newPackageBuffer,
        permissionsSizeBuffer,
        permissions,
        metadataSizeBuffer,
        newMetadataBuffer,
        endBuffer
    );

    return Buffer.from(newMashup).toString('base64');
};

type PackageComponents = {
    version: Uint8Array;
    packageOPC: Uint8Array;
    permissionsSize: number;
    permissions: Uint8Array;
    metadata: Uint8Array;
    endBuffer: Uint8Array;
};

export const getPackageComponents = (base64Str: string): PackageComponents => {
    const buffer = Buffer.from(base64Str,'base64');

    const version = buffer.slice(0,4);

    const packageSize = buffer.readInt32LE(4);
    const packageOPC = new Uint8Array(buffer.slice(8, 8 + packageSize));

    const permissionsSize = buffer.readInt32LE(8 + packageSize);
    const permissions = new Uint8Array(buffer.slice(12 + packageSize, 12 + packageSize + permissionsSize));

    const metadataSize = buffer.readInt32LE(12 + packageSize + permissionsSize);
    const metadata = new Uint8Array(buffer.slice(16 + packageSize + permissionsSize, 16 + packageSize + permissionsSize + metadataSize));

    const endBuffer = new Uint8Array(buffer.slice(16 + packageSize + permissionsSize + metadataSize))

    return {
        version,
        packageOPC,
        permissionsSize,
        permissions,
        metadata,
        endBuffer,
    };
};

const editSingleQueryPackage = async (packageOPC: ArrayBuffer, queryMashupDoc: string): Promise<Uint8Array> => {
    const packageZip: JSZip = await JSZip.loadAsync(packageOPC);
    setSection1m(queryMashupDoc, packageZip);

    return await packageZip.generateAsync({ type: uint8ArrayType });
};

const setSection1m = (queryMashupDoc: string, zip: JSZip): void => {
    if (!zip.file(section1mPath)?.async(textResultType)) {
        throw new Error(Errors.formulaSectionNotFound);
    }
    const newSection1m: string = queryMashupDoc;

    zip.file(section1mPath, newSection1m, {
        compression: emptyValue,
    });
};

export const editSingleQueryMetadata = (metadataArray: Uint8Array, metadata: Metadata): Uint8Array => {
    
    const dataView = new DataView(metadataArray.buffer, metadataArray.byteOffset, metadataArray.byteLength);
    const metadataVersion = metadataArray.slice(0, 4);
    const metadataXmlSize = dataView.getInt32(4, true);
    const metadataXml: Uint8Array = metadataArray.slice(8, 8 + metadataXmlSize);
    const endBuffer: Uint8Array = metadataArray.slice(8+metadataXmlSize);

    //parse metdataXml
    const textDecoder: TextDecoder = new TextDecoder();
    const metadataString: string = textDecoder.decode(metadataXml);
    const parser: DOMParser = new DOMParser();
    const serializer: XMLSerializer = new XMLSerializer();
    const parsedMetadata: Document = parser.parseFromString(metadataString, xmlTextResultType);

    // Update InfoPaths to new QueryName
    const itemPaths: HTMLCollectionOf<Element> = parsedMetadata.getElementsByTagName(element.itemPath);
    if (itemPaths && itemPaths.length) {
        for (let i = 0; i < itemPaths.length; i++) {
            const itemPath: Element = itemPaths[i];
            const content: string = itemPath.textContent as string;
            if (content.includes(section1PathPrefix)) {
                const strArr: string[] = content.split(divider);
                strArr[1] = encodeURIComponent(metadata.queryName);
                const newContent: string = strArr.join(divider);
                itemPath.textContent = newContent;
            }
        }
    }

    const entries = parsedMetadata.getElementsByTagName(element.entry);
    if (entries && entries.length) {
        for (let i = 0; i < entries.length; i++) {
            const entry: Element = entries[i];
            const entryAttributesArr: Attr[] = Array.from(entry.attributes);
            const entryProp: Attr | undefined = entryAttributesArr.find((prop) => {
                return prop?.name === elementAttributes.type;
            });
            if (entryProp?.nodeValue == elementAttributes.resultType) {
                entry.setAttribute(elementAttributes.value, elementAttributesValues.tableResultType());
            }

            if (entryProp?.nodeValue == elementAttributes.fillLastUpdated) {
                const nowTime: string = new Date().toISOString();
                entry.setAttribute(elementAttributes.value, (elementAttributes.day + nowTime).replace(/Z/, "0000Z"));
            }
        }
    }

    // Convert new metadataXml to Uint8Array
    const newMetadataString: string = serializer.serializeToString(parsedMetadata);
    const encoder: TextEncoder = new TextEncoder();
    const newMetadataXml: Uint8Array = encoder.encode(newMetadataString);
    const newMetadataXmlSize: Uint8Array = arrayUtils.getInt32Buffer(newMetadataXml.byteLength);
    const newMetadataArray: Uint8Array = arrayUtils.concatArrays(metadataVersion, newMetadataXmlSize, newMetadataXml, endBuffer);

    return newMetadataArray;
};
