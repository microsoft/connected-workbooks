// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import JSZip from "jszip";
import { section1mPath, defaults } from "./constants";
import { arrayUtils } from "./utils";
import { Metadata } from "./types";
import { generateSingleQueryMashup } from "./generators";

export default class MashupHandler {
    async ReplaceSingleQuery(base64Str: string, queryName:string, query: string): Promise<string> {
        const { version, packageOPC, permissionsSize, permissions, metadata, endBuffer } = this.getPackageComponents(base64Str);
        const newPackageBuffer = await this.editSingleQueryPackage(packageOPC, queryName, query);
        const packageSizeBuffer = arrayUtils.getInt32Buffer(newPackageBuffer.byteLength);
        const permissionsSizeBuffer = arrayUtils.getInt32Buffer(permissionsSize);
        const newMetadataBuffer = this.editSingleQueryMetadata(metadata, { queryName });
        const metadataSizeBuffer = arrayUtils.getInt32Buffer(newMetadataBuffer.byteLength);
        const newMashup = arrayUtils.concatArrays(version, packageSizeBuffer, newPackageBuffer, permissionsSizeBuffer, permissions, metadataSizeBuffer, newMetadataBuffer, endBuffer);
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


    protected editSingleQueryMetadata = (metadataArray: Uint8Array, metadata: Metadata) => {
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
                    itemPath.innerHTML = newContent;
                    //append child
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
                    if (entryProp?.nodeValue == "FillLastUpdated") {
                        const nowTime = new Date().toISOString();
                        entry.setAttribute("Value", ("d" + nowTime).replace(/Z/, '0000Z'));
                    }   
                }
            }
        
        // Convert new metadataXml to Uint8Array
        const newMetadataString = serializer.serializeToString(parsedMetadata);
        const util = require('util');
        const encoder = new util.TextEncoder();
        const newMetadataXml = encoder.encode(newMetadataString);
        const newMetadataXmlSize = arrayUtils.getInt32Buffer(newMetadataXml.byteLength);
        
        const newMetadataArray = arrayUtils.concatArrays(metadataVersion, newMetadataXmlSize, newMetadataXml, endBuffer);
        return newMetadataArray;
    };
}


function uintToString(uintArray: Uint8Array) {
    const util= require('util');
    const textDecoder = new util.TextDecoder();
    const encodedString = textDecoder.decode(uintArray);
    return encodedString;
 }

