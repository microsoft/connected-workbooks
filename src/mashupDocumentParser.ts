// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import JSZip from "jszip";
import { section1mPath, defaults } from "./constants";
import { arrayUtils } from "./utils";
import { Metadata } from "./types";
import { generateSingleQueryMashup } from "./generators";

export default class MashupHandler {
    async ReplaceSingleQuery(base64Str: string, queryName: string, query: string): Promise<string> {
        const { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } = this.getPackageComponents(base64Str);
        const newPackageBuffer = await this.editSingleQueryPackage(packageOPC, queryName, query);
        const packageSizeBuffer = arrayUtils.getInt32Buffer(newPackageBuffer.byteLength);
        const permissionsSizeBuffer = arrayUtils.getInt32Buffer(permissionsSize);
        const newMetadataBuffer = this.editSingleQueryMetadata(metadata, { queryName });
        const metadataSizeBuffer = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
        const newMashup = arrayUtils.concatArrays(version, packageSizeBuffer, newPackageBuffer, permissionsSizeBuffer, permissions, metadataSizeBuffer, newMetadataBuffer, endBuffer);
        return base64.fromByteArray(newMashup);
    }
    async AddConnectionOnlyQuery(base64Str: string, queryName: string): Promise<string> {
        var { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } = this.getPackageComponents(base64Str);
        const packageSizeBuffer = arrayUtils.getInt32Buffer(packageOPC.byteLength);
        const permissionsSizeBuffer = arrayUtils.getInt32Buffer(permissionsSize);
        const newMetadataBuffer = this.addConnectionOnlyQueryMetadata(metadata, queryName);
        const metadataSizeBuffer = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
        const newMashup = arrayUtils.concatArrays(version, packageSizeBuffer, packageOPC, permissionsSizeBuffer, permissions, metadataSizeBuffer, newMetadataBuffer, endBuffer);
        return base64.fromByteArray(newMashup);
    }

    private getPackageComponents(base64Str: string) {
        const buffer = base64.toByteArray(base64Str).buffer;
        const mashupArray = new arrayUtils.ArrayReader(buffer);
        const version = mashupArray.getBytes(4);
        const packageSize = mashupArray.getInt32();
        const packageOPC = mashupArray.getBytes(packageSize);
        const permissionsSize = mashupArray.getInt32();
        const permissions = mashupArray.getBytes(permissionsSize);
        const metadataSize = mashupArray.getInt32();
        const metadata = mashupArray.getBytes(metadataSize);
        const endBuffer = mashupArray.getBytes();

         return {
            version,
            packageOPC,
            permissionsSize,
            permissions,
            metadata,
            endBuffer
        };
    }

    private async editSingleQueryPackage(packageOPC: ArrayBuffer, queryName: string, query: string) {
        const packageZip = await JSZip.loadAsync(packageOPC);
        this.getSection1m(packageZip);
        this.setSection1m(queryName, query, packageZip);

        return await packageZip.generateAsync({ type: "uint8array" });
    }

    private setSection1m = (queryName: string, query: string, zip: JSZip): void => {
        const newSection1m = generateSingleQueryMashup(queryName, query);

        zip.file(section1mPath, newSection1m, {
            compression: "",
        });
    };

    private getSection1m = async (zip: JSZip): Promise<string> => {
        const section1m = zip.file(section1mPath)?.async("text");
        if (!section1m) {
            throw new Error("Formula section wasn't found in template");
        }

        return section1m;
    };

    
    private editSingleQueryMetadata = (metadataArray: Uint8Array, metadata: Metadata) => {
        //extract metadataXml
        const mashupArray = new arrayUtils.ArrayReader(metadataArray.buffer);
        const metadataVersion = mashupArray.getBytes(4);
        const metadataXmlSize = mashupArray.getInt32();
        const metadataXml = mashupArray.getBytes(metadataXmlSize);
        const endBuffer = mashupArray.getBytes();

        //parse metdataXml
        const textDecoder = new TextDecoder();
        const metadataString = textDecoder.decode(metadataXml);
        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const parsedMetadata = parser.parseFromString(metadataString, "text/xml");

        // Update InfoPaths to new QueryName
        const itemPaths = parsedMetadata.getElementsByTagName("ItemPath");
        if (itemPaths && itemPaths.length) {
            for (let i = 0; i < itemPaths.length; i++) {
                const itemPath = itemPaths[i];
                const content = itemPath.innerHTML;
                if (content.includes("Section1/")) {
                    const strArr = content.split("/");
                    strArr[1] = metadata.queryName;
                    const newContent = strArr.join("/");
                    itemPath.textContent = newContent;
                    }    
                }
            }

        const entries = parsedMetadata.getElementsByTagName("Entry");
            if (entries && entries.length) {
                for (let i = 0; i < entries.length; i++) {
                    const entry = entries[i];
                    const entryAttributes = entry.attributes;
                    const entryAttributesArr = [...entryAttributes]; 
                    const entryProp = entryAttributesArr.find((prop) => {
                    return prop?.name === "Type"});
                    if (entryProp?.nodeValue == "RelationshipInfoContainer") {
                        const newValue = entry.getAttribute("Value")?.replace(/Query1/g, metadata.queryName);
                        if (newValue) {
                            entry.setAttribute("Value", newValue);
                        }
                    }
                    if (entryProp?.nodeValue == "ResultType") {
                        entry.setAttribute("Value", "sTable");
                    }
                    if (entryProp?.nodeValue == "FillColumnNames") {
                        const oldValue = entry.getAttribute("Value");
                        if (oldValue) {
                            entry.setAttribute("Value", oldValue.replace(defaults.queryName, metadata.queryName));
                        }    
                    }
                    if (entryProp?.nodeValue == "FillTarget") {
                        const oldValue = entry.getAttribute("Value");
                        if (oldValue) {
                            entry.setAttribute("Value", oldValue.replace(defaults.queryName, metadata.queryName));
                        }    
                    }
                    if (entryProp?.nodeValue == "FillLastUpdated") {
                        const nowTime = new Date().toISOString();
                        entry.setAttribute("Value", ("d" + nowTime).replace(/Z/, '0000Z'));
                    }   
                }
            }

        // Convert new metadataXml to Uint8Array
        const newMetadataString = serializer.serializeToString(parsedMetadata);
        const encoder = new TextEncoder();
        const newMetadataXml = encoder.encode(newMetadataString);
        const newMetadataXmlSize = arrayUtils.getInt32Buffer(newMetadataXml.byteLength);

        const newMetadataArray = arrayUtils.concatArrays(metadataVersion, newMetadataXmlSize, newMetadataXml, endBuffer);
        return newMetadataArray;
    };

    private addConnectionOnlyQueryMetadata = (metadataArray: Uint8Array, queryName: string) => {
        //extract metadataXml
        const mashupArray = new arrayUtils.ArrayReader(metadataArray.buffer);
        const metadataVersion = mashupArray.getBytes(4);
        const metadataXmlSize = mashupArray.getInt32();
        const metadataXml = mashupArray.getBytes(metadataXmlSize);
        const endBuffer = mashupArray.getBytes();

        //parse metadataXml
        const metadataString = uintToString(metadataXml);
        const newMetadataString = this.updateConnectionOnlyMetadataStr(metadataString, queryName);

        const encoder = new TextEncoder();
        const newMetadataXml = encoder.encode(newMetadataString);
        const newMetadataXmlSize = arrayUtils.getInt32Buffer(newMetadataXml.byteLength);
        const newMetadataArray = arrayUtils.concatArrays(
            metadataVersion,
            newMetadataXmlSize,
            newMetadataXml,
            endBuffer
        );
        return newMetadataArray;
    };

    private updateConnectionOnlyMetadataStr = (metadataString: string, queryName: string) => {
        const parser = new DOMParser();
        const metadataDoc = parser.parseFromString(metadataString, "text/xml");     
        this.createStableEntriesItem(metadataDoc, queryName);
        this.createSourceItem(metadataDoc, queryName);

        const serializer = new XMLSerializer();
        const newMetadataString = serializer.serializeToString(metadataDoc);
        return newMetadataString;
    };

    private createSourceItem = (metadataDoc: Document, queryName: string) => {
        const items = metadataDoc.getElementsByTagName("Items")[0];
        const newItemSource = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Item");
        const newItemLocation = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemLocation");
        const newItemType = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemType");
        newItemType.textContent = "Formula";
        const newItemPath = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemPath");
        newItemPath.textContent = `Section1/${queryName}/Source`;
        newItemLocation.appendChild(newItemType);
        newItemLocation.appendChild(newItemPath);
        newItemSource.appendChild(newItemLocation);
        items.appendChild(newItemSource);
    };

    private createStableEntriesItem = (metadataDoc: Document, queryName: string) => {
        const items = metadataDoc.getElementsByTagName("Items")[0];
        const newItem = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Item");
        const newItemLocation = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemLocation");
        const newItemType = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemType");
        newItemType.textContent = "Formula";
        const newItemPath = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemPath");
        newItemPath.textContent = `Section1/${queryName}`;
        newItemLocation.appendChild(newItemType);
        newItemLocation.appendChild(newItemPath); 
        newItem.appendChild(newItemLocation);
        items.append(newItem);
        let stableEntries = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "StableEntries");
        stableEntries = this.createConnectionOnlyEntries(metadataDoc, stableEntries);
        newItem.appendChild(stableEntries);
    };

    private createConnectionOnlyEntries = (metadataDoc: Document, stableEntries: Element) => {
        const IsPrivate = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        IsPrivate.setAttribute("Type", "IsPrivate");
        IsPrivate.setAttribute("Value", "l0");
        stableEntries.appendChild(IsPrivate);
        const FillEnabled = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillEnabled.setAttribute("Type", "FillEnabled");
        FillEnabled.setAttribute("Value", "l0");
        stableEntries.appendChild(FillEnabled);
        const FillObjectType = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillObjectType.setAttribute("Type", "FillObjectType");
        FillObjectType.setAttribute("Value", "sConnectionOnly");
        stableEntries.appendChild(FillObjectType);
        const FillToDataModelEnabled = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillToDataModelEnabled.setAttribute("Type", "FillToDataModelEnabled");
        FillToDataModelEnabled.setAttribute("Value", "l0");
        const FillLastUpdated = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillLastUpdated.setAttribute("Type", "FillLastUpdated");
        const nowTime = new Date().toISOString();
        FillLastUpdated.setAttribute("Value", ("d" + nowTime).replace(/Z/, "0000Z"));
        stableEntries.appendChild(FillLastUpdated);
        const ResultType = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        ResultType.setAttribute("Type", "ResultType");
        ResultType.setAttribute("Value", "sTable");
        stableEntries.appendChild(ResultType);
        return stableEntries;
    };
}

function uintToString(uintArray: Uint8Array) {
    var encodedString = new TextDecoder("utf-8").decode(uintArray);
    return encodedString;
}
