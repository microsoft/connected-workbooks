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

    return base64.fromByteArray(newMashup);
};

export const addConnectionOnlyQuery = async (base64Str: string, connectionOnlyQueryNames: string[]): Promise<string> => {
        var { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } = getPackageComponents(base64Str);
        const packageSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(packageOPC.byteLength);
        const permissionsSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(permissionsSize);
        const newMetadataBuffer: Uint8Array = addConnectionOnlyQueryMetadata(metadata, connectionOnlyQueryNames);
        const metadataSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
        const newMashup: Uint8Array = arrayUtils.concatArrays(version, packageSizeBuffer, packageOPC, permissionsSizeBuffer, permissions, metadataSizeBuffer, newMetadataBuffer, endBuffer);
        
        return base64.fromByteArray(newMashup);
    }

type PackageComponents = {
    version: Uint8Array;
    packageOPC: Uint8Array;
    permissionsSize: number;
    permissions: Uint8Array;
    metadata: Uint8Array;
    endBuffer: Uint8Array;
};

export const getPackageComponents = (base64Str: string): PackageComponents => {
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

const editSingleQueryPackage = async (packageOPC: ArrayBuffer, queryMashupDoc: string): Promise<Uint8Array> => {
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

export const editSingleQueryMetadata = (metadataArray: Uint8Array, metadata: Metadata): Uint8Array => {
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
            const entryAttributes: NamedNodeMap = entry.attributes;
            const entryAttributesArr: Attr[] = [...entryAttributes];
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

export const addConnectionOnlyQueryMetadata = (metadataArray: Uint8Array, connectionOnlyQueryNames: string[]) => {
        // extract metadataXml
        const mashupArray: ArrayReader = new arrayUtils.ArrayReader(metadataArray.buffer);
        const metadataVersion: Uint8Array = mashupArray.getBytes(4);
        const metadataXmlSize: number = mashupArray.getInt32();
        const metadataXml: Uint8Array = mashupArray.getBytes(metadataXmlSize);
        const endBuffer: Uint8Array = mashupArray.getBytes();

        // parse metadataXml
        const metadataString: string = new TextDecoder("utf-8").decode(metadataXml);
        const newMetadataString: string = updateConnectionOnlyMetadataStr(metadataString, connectionOnlyQueryNames);
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

    const updateConnectionOnlyMetadataStr = (metadataString: string, connectionOnlyQueryNames: string[]) => {
        const parser: DOMParser = new DOMParser();
        let updatedMetdataString: string = metadataString;         
        connectionOnlyQueryNames.forEach((queryName: string) => {
            const metadataDoc: Document = parser.parseFromString(updatedMetdataString, xmlTextResultType);
            const items: Element = metadataDoc.getElementsByTagName(element.items)[0];
            const stableEntriesItem: Element = createStableEntriesItem(metadataDoc, queryName);
            items.appendChild(stableEntriesItem);
            const sourceItem: Element = createSourceItem(metadataDoc, queryName);
            items.appendChild(sourceItem);
            const serializer: XMLSerializer = new XMLSerializer();
            updatedMetdataString = serializer.serializeToString(metadataDoc);
        });    

        return updatedMetdataString;
    };

    const createSourceItem = (metadataDoc: Document, queryName: string) => {
        const newItemSource: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.item);
        const newItemLocation: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.itemLocation);
        const newItemType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.itemType);
        newItemType.textContent = "Formula";
        const newItemPath: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.itemPath);
        newItemPath.textContent = `Section1/${queryName}/Source`;
        newItemLocation.appendChild(newItemType);
        newItemLocation.appendChild(newItemPath);
        newItemSource.appendChild(newItemLocation);
        
        return newItemSource;
    };

    const createStableEntriesItem = (metadataDoc: Document, queryName: string) => {
        const newItem: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.item);
        const newItemLocation: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.itemLocation);
        const newItemType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.itemType);
        newItemType.textContent = "Formula";
        const newItemPath: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.itemPath);
        newItemPath.textContent = `Section1/${queryName}`;
        newItemLocation.appendChild(newItemType);
        newItemLocation.appendChild(newItemPath); 
        newItem.appendChild(newItemLocation);
        const stableEntries: Element = createConnectionOnlyEntries(metadataDoc);
        newItem.appendChild(stableEntries);
        
        return newItem;
    };

    const createConnectionOnlyEntries = (metadataDoc: Document) => {
        const stableEntries: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.stableEntries);
        
        const IsPrivate: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        IsPrivate.setAttribute(elementAttributes.type, elementAttributes.isPrivate);
        IsPrivate.setAttribute(elementAttributes.value, "l0");
        
        stableEntries.appendChild(IsPrivate);
        const FillEnabled: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        FillEnabled.setAttribute(elementAttributes.type, elementAttributes.fillEnabled);
        FillEnabled.setAttribute(elementAttributes.value, "l0");
        stableEntries.appendChild(FillEnabled);
        
        const FillObjectType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        FillObjectType.setAttribute(elementAttributes.type, elementAttributes.fillObjectType);
        FillObjectType.setAttribute(elementAttributes.value, elementAttributesValues.connectionOnlyResultType());
        stableEntries.appendChild(FillObjectType);
        
        const FillToDataModelEnabled: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        FillToDataModelEnabled.setAttribute(elementAttributes.type, elementAttributes.fillToDataModelEnabled);
        FillToDataModelEnabled.setAttribute(elementAttributes.value, "l0");
        stableEntries.appendChild(FillToDataModelEnabled);
        
        const ResultType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        ResultType.setAttribute(elementAttributes.type, elementAttributes.resultType);
        ResultType.setAttribute(elementAttributes.value, elementAttributesValues.tableResultType());
        stableEntries.appendChild(ResultType);
        
        const BufferNextRefresh: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        BufferNextRefresh.setAttribute(elementAttributes.type, elementAttributes.bufferNextRefresh);
        BufferNextRefresh.setAttribute(elementAttributes.value, "l1");
        stableEntries.appendChild(BufferNextRefresh);
        
        const FilledCompleteResultToWorksheet: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        FilledCompleteResultToWorksheet.setAttribute(elementAttributes.type, elementAttributes.filledCompleteResultToWorksheet);
        FilledCompleteResultToWorksheet.setAttribute(elementAttributes.value, "l0");
        stableEntries.appendChild(FilledCompleteResultToWorksheet);
             
        const AddedToDataModel: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        AddedToDataModel.setAttribute(elementAttributes.type, elementAttributes.addedToDataModel);
        AddedToDataModel.setAttribute(elementAttributes.value, "l0");
        stableEntries.appendChild(AddedToDataModel);

        const FillErrorCode: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        FillErrorCode.setAttribute(elementAttributes.type, elementAttributes.fillErrorCode);
        FillErrorCode.setAttribute(elementAttributes.value, elementAttributesValues.fillErrorCodeUnknown());
        stableEntries.appendChild(FillErrorCode);

        const FillLastUpdated: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        FillLastUpdated.setAttribute(elementAttributes.type, elementAttributes.fillLastUpdated);
        const nowTime: string = new Date().toISOString();
        FillLastUpdated.setAttribute(elementAttributes.value, (elementAttributes.day + nowTime).replace(/Z/, "0000Z"));
        stableEntries.appendChild(FillLastUpdated);
        
        const FillStatus: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, element.entry);
        FillStatus.setAttribute(elementAttributes.type, elementAttributes.fillStatus);
        FillStatus.setAttribute(elementAttributes.value, elementAttributesValues.fillStatusComplete());
        stableEntries.appendChild(FillStatus);
        
        return stableEntries;
    };