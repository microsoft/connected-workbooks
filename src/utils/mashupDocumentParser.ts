// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import JSZip from "jszip";
import {
    section1mPath,
    defaults,
    uint8ArrayType,
    emptyValue,
    textResultType,
    formulaSectionNotFoundErr,
    xmlTextResultType,
    element,
    section1PathPrefix,
    divider,
    elementAttributes,
    elementAttributesValues,
} from "./constants";
import { arrayUtils } from ".";
import { Metadata } from "../types";
import { ArrayReader } from "./arrayUtils";

export const replaceSingleQuery = async (
    base64Str: string,
    queryName: string,
    queryMashupDoc: string
): Promise<string> => {
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

    return base64.fromByteArray(newMashup);
};

export const getPackageComponents = (base64Str: string) => {
    const buffer: ArrayBufferLike = base64.toByteArray(base64Str).buffer;
    const mashupArray: ArrayReader = new arrayUtils.ArrayReader(buffer);
    const version: Uint8Array = mashupArray.getBytes(4);
    const packageSize: number = mashupArray.getInt32();
    const packageOPC: Uint8Array = mashupArray.getBytes(packageSize);
    const permissionsSize: number = mashupArray.getInt32();
    const permissions: Uint8Array = mashupArray.getBytes(permissionsSize);
    const metadataSize: number = mashupArray.getInt32();
    const metadata: Uint8Array = mashupArray.getBytes(metadataSize);
    const endBuffer: Uint8Array = mashupArray.getBytes();

    return {
        version,
        packageOPC,
        permissionsSize,
        permissions,
        metadata,
        endBuffer,
    };
};

export const editSingleQueryPackage = async (packageOPC: ArrayBuffer, queryMashupDoc: string) => {
    const packageZip: JSZip = await JSZip.loadAsync(packageOPC);
    setSection1m(queryMashupDoc, packageZip);

    return await packageZip.generateAsync({ type: uint8ArrayType });
};

const setSection1m = (queryMashupDoc: string, zip: JSZip): void => {
    if (!zip.file(section1mPath)?.async(textResultType)) {
        throw new Error(formulaSectionNotFoundErr);
    }
    const newSection1m: string = queryMashupDoc;

    zip.file(section1mPath, newSection1m, {
        compression: emptyValue,
    });
};

const editSingleQueryMetadata = (metadataArray: Uint8Array, metadata: Metadata) => {
    //extract metadataXml
    const mashupArray: ArrayReader = new arrayUtils.ArrayReader(metadataArray.buffer);
    const metadataVersion: Uint8Array = mashupArray.getBytes(4);
    const metadataXmlSize: number = mashupArray.getInt32();
    const metadataXml: Uint8Array = mashupArray.getBytes(metadataXmlSize);
    const endBuffer: Uint8Array = mashupArray.getBytes();

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
            const content: string = itemPath.innerHTML;
            if (content.includes(section1PathPrefix)) {
                const strArr: string[] = content.split(divider);
                strArr[1] = metadata.queryName;
                const newContent: string = strArr.join(divider);
                itemPath.textContent = newContent;
            }
        }
    }

    const entries = parsedMetadata.getElementsByTagName(element.entry);
    if (entries && entries.length) {
        for (let i = 0; i < entries.length; i++) {
            const entry: Element = entries[i];
            const entryAttributes: NamedNodeMap = entry.attributes;
            const entryAttributesArr: Attr[] = [...entryAttributes];
            const entryProp: Attr | undefined = entryAttributesArr.find((prop) => {
                return prop?.name === elementAttributes.type;
            });
            if (entryProp?.nodeValue == elementAttributes.relationshipInfo) {
                const newValue: string | undefined = entry
                    .getAttribute(elementAttributes.value)
                    ?.replace(/Query1/g, metadata.queryName);
                if (newValue) {
                    entry.setAttribute(elementAttributes.value, newValue);
                }
            }
            if (entryProp?.nodeValue == elementAttributes.resultType) {
                entry.setAttribute(elementAttributes.value, elementAttributesValues.tableResultType());
            }

            if (entryProp?.nodeValue == elementAttributes.fillColumnNames) {
                const oldValue: string | null = entry.getAttribute(elementAttributes.value);
                if (oldValue) {
                    entry.setAttribute(
                        elementAttributes.value,
                        oldValue.replace(defaults.queryName, metadata.queryName)
                    );
                }
            }

            if (entryProp?.nodeValue == elementAttributes.fillTarget) {
                const oldValue: string | null = entry.getAttribute(elementAttributes.value);
                if (oldValue) {
                    entry.setAttribute(
                        elementAttributes.value,
                        oldValue.replace(defaults.queryName, metadata.queryName)
                    );
                }
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
    const newMetadataArray: Uint8Array = arrayUtils.concatArrays(
        metadataVersion,
        newMetadataXmlSize,
        newMetadataXml,
        endBuffer
    );

    return newMetadataArray;
};
