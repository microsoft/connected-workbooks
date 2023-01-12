// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import JSZip from "jszip";
import { section1mPath, defaults } from "./constants";
import { arrayUtils } from "./utils";
import { Metadata } from "./types";
import { generateSingleQueryMashup, generateConnectionOnlyQueryMashup } from "./generators";

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

     async AddConnectionOnlyQuery(base64Str: string, queryName: string, query: string, connectionOnlyName: string, connectionOnlyQuery: string): Promise<string> {
        const { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } = this.getPackageComponents(base64Str);
        const newPackageBuffer = await this.addConnectionOnlyQueryPackage(packageOPC, queryName, query, connectionOnlyName, connectionOnlyQuery);
        const packageSizeBuffer = arrayUtils.getInt32Buffer(newPackageBuffer.byteLength);
        const permissionsSizeBuffer = arrayUtils.getInt32Buffer(permissionsSize);
        const newMetadataBuffer = this.addConnectionOnlyMetadata(metadata, connectionOnlyName);
        const metadataSizeBuffer = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
        const newMashup = arrayUtils.concatArrays(version, packageSizeBuffer, newPackageBuffer, permissionsSizeBuffer, 
            permissions, metadataSizeBuffer, newMetadataBuffer, endBuffer);
        return base64.fromByteArray(newMashup);
    }

       private async addConnectionOnlyQueryPackage(packageOPC: ArrayBuffer, queryName: string, query: string, connectionOnlyName: string, connectionOnlyQuery: string) {
        const packageZip = await JSZip.loadAsync(packageOPC);
        this.getSection1m(packageZip);
        this.updateSection1mWithConnectionOnly(queryName, query, connectionOnlyName, connectionOnlyQuery, packageZip);

        return await packageZip.generateAsync({ type: "uint8array" });
    }

    private updateSection1mWithConnectionOnly = async (queryName: string, query: string, connectionOnlyName: string, connectionOnlyQuery: string, zip: JSZip): Promise<void> => {
        const newSection1m = generateConnectionOnlyQueryMashup(queryName, query, connectionOnlyName, connectionOnlyQuery);
        zip.file(section1mPath, newSection1m, {
            compression: "",
        });
    };

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

    private addConnectionOnlyMetadata = (metadataArray: Uint8Array, queryName: string) => {
        //extract metadataXml
        const mashupArray = new arrayUtils.ArrayReader(metadataArray.buffer);
        const metadataVersion = mashupArray.getBytes(4);
        const metadataXmlSize = mashupArray.getInt32();
        const metadataXml = mashupArray.getBytes(metadataXmlSize);
        const endBuffer = mashupArray.getBytes();
        
        //parse metdataXml
        const metadataString = uintToString(metadataXml);
        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const metadataDoc = parser.parseFromString(metadataString, "text/xml");
        
        // Add new Item to metadataXml
        const oldItem = metadataDoc.getElementsByTagName("Item")[1];
        const oldPathItem = metadataDoc.getElementsByTagName("Item")[2];
        const items = metadataDoc.getElementsByTagName("Items")[0];
        const newItem = oldItem.cloneNode(false);
        const itemLocation = metadataDoc.getElementsByTagName("ItemLocation")[1].cloneNode(true);
        newItem.appendChild(itemLocation);
        const stableEntries = metadataDoc.getElementsByTagName("StableEntries")[1].cloneNode(true);
        newItem.appendChild(stableEntries);
        items.appendChild(newItem);
        const entries = metadataDoc.getElementsByTagName("Entry");
        if (entries && entries.length) {
            for (let i = 0; i < entries.length; i++) {
                const entry = entries[i];
                if (entry.parentElement === stableEntries) {
                    const entryAttributesArr = [...entry.attributes]; 
                    const entryProp = entryAttributesArr.find((prop) => {
                        return prop?.name === "Type"});
                    if (entryProp?.nodeValue == "ResultType") {
                        entry.setAttribute("Value", "sTable");
                    }
                    if (entryProp?.nodeValue == "FillObjectType") {
                        entry.setAttribute("Value", "sConnectionOnly");
                    }
                    if (entryProp?.nodeValue == "FillEnabled") {
                        entry.setAttribute("Value", "l0");
                    }
                    if (entryProp?.nodeValue == "FillLastUpdated") {
                        const nowTime = new Date().toISOString();
                        entry.setAttribute("Value", ("d" + nowTime).replace(/Z/, '0000Z'));
                    }
                    if (entryProp?.nodeValue == "FillColumnNames") {
                        stableEntries.removeChild(entry);
                    }
                    if (entryProp?.nodeValue == "FillTarget") {
                        stableEntries.removeChild(entry);
                    }
                }                  
            }
        }
        const newPathItem = oldPathItem.cloneNode(true);
        items.appendChild(newPathItem);
        const itemPaths = metadataDoc.getElementsByTagName("ItemPath");
        if (itemPaths && itemPaths.length) {
            for (let i = 0; i < itemPaths.length; i++) {
                const itemPath = itemPaths[i];
                if ((itemPath.parentElement?.parentElement === newPathItem) || (itemPath.parentElement === itemLocation)) {
                    const content = itemPath.innerHTML;
                    if (content.includes("Section1/")) {
                        const strArr = content.split("/");
                        strArr[1] = queryName;
                        const newContent = strArr.join("/");
                        itemPath.textContent = newContent;
                    }    
                } 
            }
        }
        // Convert new metadataXml to Uint8Array
        const newMetadataString = serializer.serializeToString(metadataDoc);
        const encoder = new TextEncoder();
        const newMetadataXml = encoder.encode(newMetadataString);
        const newMetadataXmlSize = arrayUtils.getInt32Buffer(newMetadataXml.byteLength); 
        const newMetadataArray = arrayUtils.concatArrays(metadataVersion, newMetadataXmlSize, newMetadataXml, endBuffer);
        return newMetadataArray;
    };
}

function uintToString(uintArray: Uint8Array) {
    var encodedString = new TextDecoder("utf-8").decode(uintArray);
    return encodedString;
 }
