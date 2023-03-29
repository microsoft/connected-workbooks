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
        const newPackageBuffer: Uint8Array = await this.editSingleQueryPackage(packageOPC, queryName, query);
        const packageSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(newPackageBuffer.byteLength);
        const permissionsSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(permissionsSize);
        const newMetadataBuffer: Uint8Array = this.editSingleQueryMetadata(metadata, { queryName });
        const metadataSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
        const newMashup: Uint8Array = arrayUtils.concatArrays(version, packageSizeBuffer, newPackageBuffer, permissionsSizeBuffer, permissions, metadataSizeBuffer, newMetadataBuffer, endBuffer);
        return base64.fromByteArray(newMashup);
    }

    async AddConnectionOnlyQuery(base64Str: string, queryName: string): Promise<string> {
        var { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } = this.getPackageComponents(base64Str);
        const packageSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(packageOPC.byteLength);
        const permissionsSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(permissionsSize);
        const newMetadataBuffer: Uint8Array = this.addConnectionOnlyQueryMetadata(metadata, queryName);
        const metadataSizeBuffer: Uint8Array = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
        const newMashup: Uint8Array = arrayUtils.concatArrays(version, packageSizeBuffer, packageOPC, permissionsSizeBuffer, permissions, metadataSizeBuffer, newMetadataBuffer, endBuffer);
        return base64.fromByteArray(newMashup);
    }

    private getPackageComponents(base64Str: string) {
        const buffer: ArrayBufferLike = base64.toByteArray(base64Str).buffer;
        const mashupArray: any = new arrayUtils.ArrayReader(buffer);
        const version: any = mashupArray.getBytes(4);
        const packageSize: any = mashupArray.getInt32();
        const packageOPC: any = mashupArray.getBytes(packageSize);
        const permissionsSize: any = mashupArray.getInt32();
        const permissions: any = mashupArray.getBytes(permissionsSize);
        const metadataSize: any = mashupArray.getInt32();
        const metadata: any = mashupArray.getBytes(metadataSize);
        const endBuffer: any = mashupArray.getBytes();

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
        const packageZip: JSZip = await JSZip.loadAsync(packageOPC);
        this.getSection1m(packageZip);
        this.setSection1m(queryName, query, packageZip);

        return await packageZip.generateAsync({ type: "uint8array" });
    }

    private setSection1m = (queryName: string, query: string, zip: JSZip): void => {
        const newSection1m: string = generateSingleQueryMashup(queryName, query);

        zip.file(section1mPath, newSection1m, {
            compression: "",
        });
    };

    private getSection1m = async (zip: JSZip): Promise<string> => {
        const section1m: Promise<string> | undefined = zip.file(section1mPath)?.async("text");
        if (!section1m) {
            throw new Error("Formula section wasn't found in template");
        }

        return section1m;
    };

    
    private editSingleQueryMetadata = (metadataArray: Uint8Array, metadata: Metadata) => {
        //extract metadataXml
        const mashupArray: any = new arrayUtils.ArrayReader(metadataArray.buffer);
        const metadataVersion: any = mashupArray.getBytes(4);
        const metadataXmlSize: any = mashupArray.getInt32();
        const metadataXml: any = mashupArray.getBytes(metadataXmlSize);
        const endBuffer: any = mashupArray.getBytes();

        //parse metdataXml
        const textDecoder: TextDecoder = new TextDecoder();
        const metadataString: string = textDecoder.decode(metadataXml);
        const parser: DOMParser = new DOMParser();
        const serializer: XMLSerializer = new XMLSerializer();
        const parsedMetadata: Document = parser.parseFromString(metadataString, "text/xml");

        // Update InfoPaths to new QueryName
        const itemPaths: HTMLCollectionOf<Element> = parsedMetadata.getElementsByTagName("ItemPath");
        if (itemPaths && itemPaths.length) {
            for (let i = 0; i < itemPaths.length; i++) {
                const itemPath: Element = itemPaths[i];
                const content: string = itemPath.innerHTML;
                if (content.includes("Section1/")) {
                    const strArr: string[] = content.split("/");
                    strArr[1] = metadata.queryName;
                    const newContent: string = strArr.join("/");
                    itemPath.textContent = newContent;
                    }    
                }
            }

        const entries = parsedMetadata.getElementsByTagName("Entry");
            if (entries && entries.length) {
                for (let i = 0; i < entries.length; i++) {
                    const entry: Element = entries[i];
                    const entryAttributes: NamedNodeMap = entry.attributes;
                    const entryAttributesArr: Attr[] = [...entryAttributes]; 
                    const entryProp: Attr | undefined = entryAttributesArr.find((prop) => {
                    return prop?.name === "Type"});
                    if (entryProp?.nodeValue == "RelationshipInfoContainer") {
                        const newValue: string | undefined = entry.getAttribute("Value")?.replace(/Query1/g, metadata.queryName);
                        if (newValue) {
                            entry.setAttribute("Value", newValue);
                        }
                    }
                    if (entryProp?.nodeValue == "ResultType") {
                        entry.setAttribute("Value", "sTable");
                    }
                    if (entryProp?.nodeValue == "FillColumnNames") {
                        const oldValue: string | null = entry.getAttribute("Value");
                        if (oldValue) {
                            entry.setAttribute("Value", oldValue.replace(defaults.queryName, metadata.queryName));
                        }    
                    }
                    if (entryProp?.nodeValue == "FillTarget") {
                        const oldValue: string | null = entry.getAttribute("Value");
                        if (oldValue) {
                            entry.setAttribute("Value", oldValue.replace(defaults.queryName, metadata.queryName));
                        }    
                    }
                    if (entryProp?.nodeValue == "FillLastUpdated") {
                        const nowTime: string = new Date().toISOString();
                        entry.setAttribute("Value", ("d" + nowTime).replace(/Z/, '0000Z'));
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

    private addConnectionOnlyQueryMetadata = (metadataArray: Uint8Array, queryName: string) => {
        //extract metadataXml
        const mashupArray: any = new arrayUtils.ArrayReader(metadataArray.buffer);
        const metadataVersion: any = mashupArray.getBytes(4);
        const metadataXmlSize: any = mashupArray.getInt32();
        const metadataXml: any = mashupArray.getBytes(metadataXmlSize);
        const endBuffer: any = mashupArray.getBytes();

        //parse metadataXml
        const metadataString: string = uintToString(metadataXml);
        const newMetadataString: string = this.updateConnectionOnlyMetadataStr(metadataString, queryName);

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

    private updateConnectionOnlyMetadataStr = (metadataString: string, queryName: string) => {
        const parser: DOMParser = new DOMParser();
        const metadataDoc: Document = parser.parseFromString(metadataString, "text/xml");     
        const items: Element = metadataDoc.getElementsByTagName("Items")[0];
        const stableEntriesItem: Element = this.createStableEntriesItem(metadataDoc, queryName);
        items.appendChild(stableEntriesItem);
        const sourceItem: Element = this.createSourceItem(metadataDoc, queryName);
        items.appendChild(sourceItem);
        const serializer: XMLSerializer = new XMLSerializer();
        const newMetadataString: string = serializer.serializeToString(metadataDoc);
        return newMetadataString;
    };

    private createSourceItem = (metadataDoc: Document, queryName: string) => {
        const newItemSource: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Item");
        const newItemLocation: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemLocation");
        const newItemType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemType");
        newItemType.textContent = "Formula";
        const newItemPath: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemPath");
        newItemPath.textContent = `Section1/${queryName}/Source`;
        newItemLocation.appendChild(newItemType);
        newItemLocation.appendChild(newItemPath);
        newItemSource.appendChild(newItemLocation);
        return newItemSource;
    };

    private createStableEntriesItem = (metadataDoc: Document, queryName: string) => {
        const newItem: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Item");
        const newItemLocation: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemLocation");
        const newItemType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemType");
        newItemType.textContent = "Formula";
        const newItemPath: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "ItemPath");
        newItemPath.textContent = `Section1/${queryName}`;
        newItemLocation.appendChild(newItemType);
        newItemLocation.appendChild(newItemPath); 
        newItem.appendChild(newItemLocation);
        const stableEntries: Element = this.createConnectionOnlyEntries(metadataDoc);
        newItem.appendChild(stableEntries);
        return newItem;
    };

    private createConnectionOnlyEntries = (metadataDoc: Document) => {
        const stableEntries: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "StableEntries");
        const IsPrivate: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        IsPrivate.setAttribute("Type", "IsPrivate");
        IsPrivate.setAttribute("Value", "l0");
        stableEntries.appendChild(IsPrivate);
        const FillEnabled: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillEnabled.setAttribute("Type", "FillEnabled");
        FillEnabled.setAttribute("Value", "l0");
        stableEntries.appendChild(FillEnabled);
        const FillObjectType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillObjectType.setAttribute("Type", "FillObjectType");
        FillObjectType.setAttribute("Value", "sConnectionOnly");
        stableEntries.appendChild(FillObjectType);
        const FillToDataModelEnabled: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillToDataModelEnabled.setAttribute("Type", "FillToDataModelEnabled");
        FillToDataModelEnabled.setAttribute("Value", "l0");
        stableEntries.appendChild(FillToDataModelEnabled);
        const FillLastUpdated: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        FillLastUpdated.setAttribute("Type", "FillLastUpdated");
        const nowTime: string = new Date().toISOString();
        FillLastUpdated.setAttribute("Value", ("d" + nowTime).replace(/Z/, "0000Z"));
        stableEntries.appendChild(FillLastUpdated);
        const ResultType: Element = metadataDoc.createElementNS(metadataDoc.documentElement.namespaceURI, "Entry");
        ResultType.setAttribute("Type", "ResultType");
        ResultType.setAttribute("Value", "sTable");
        stableEntries.appendChild(ResultType);
        return stableEntries;
    };
}

function uintToString(uintArray: Uint8Array) {
    var encodedString: string = new TextDecoder("utf-8").decode(uintArray);
    return encodedString;
}
