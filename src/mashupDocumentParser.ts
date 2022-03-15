// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import JSZip from "jszip";
import { section1mPath, defaults } from "./constants";
import { Metadata } from "./types";
import { arrayUtils } from "./utils";
import { generateSingleQueryMashup } from "./generators";

export default class MashupHandler {
    async ReplaceSingleQuery(
        base64Str: string,
        query: string,
        queryName: string = defaults.queryName
    ): Promise<string> {
        const { packageOPC, startArray, endBuffer } =
            this.getPackageComponents(base64Str);

        const newPackageBuffer = await this.editSingleQueryPackage(
            packageOPC,
            queryName,
            query
        );
        const newEndBuffer = this.editSingleQueryMetadata(
            base64Str,
            endBuffer,
            { queryName }
        );

        const packageSizeBuffer = arrayUtils.getInt32Buffer(
            newPackageBuffer.byteLength
        );

        const newMashup = arrayUtils.concatArrays(
            startArray,
            packageSizeBuffer,
            newPackageBuffer,
            newEndBuffer
        );
        return base64.fromByteArray(newMashup);
    }

    private getPackageComponents(base64Str: string) {
        const buffer = base64.toByteArray(base64Str).buffer;
        const mashupArray = new arrayUtils.ArrayReader(buffer);
        const startArray = mashupArray.getBytes(4);
        const packageSize = mashupArray.getInt32();
        const packageOPC = mashupArray.getBytes(packageSize);
        const endBuffer = mashupArray.getBytes();

        return {
            startArray,
            packageOPC,
            endBuffer,
        };
    }

    private async editSingleQueryPackage(
        packageOPC: ArrayBuffer,
        queryName: string,
        query: string
    ) {
        const packageZip = await JSZip.loadAsync(packageOPC);
        this.getSection1m(packageZip);
        this.setSection1m(queryName, query, packageZip);

        return await packageZip.generateAsync({ type: "uint8array" });
    }

    private setSection1m = (
        queryName: string,
        query: string,
        zip: JSZip
    ): void => {
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

    private editSingleQueryMetadata = (
        base64Str: string,
        endBuffer: Uint8Array,
        metadata: Metadata
    ) => {
        const metaDataXml = this.getMetadataXml(base64Str, endBuffer);
        const newEndBuffer = this.setMetadataXml(metaDataXml, metadata);

        const encoder = new TextEncoder();
        return encoder.encode(newEndBuffer);
    };

    private getMetadataXml = (base64Str: string, endBuffer: Uint8Array) => {
        const arrayReader = new arrayUtils.ArrayReader(endBuffer);
        const decoded = base64.toByteArray(base64Str);
        let version = arrayReader.getBytes(4);
        const packageSize = arrayReader.getBytes(4);
        const permissionsSize = arrayReader.getBytes(
            4 + packageSize.byteLength
        );
        const metadataSize = arrayReader.getBytes(
            4 + permissionsSize.byteLength
        );
        const metadataVersion = arrayReader.getBytes(4);
        const metadataXmlSize = arrayReader.getBytes(4);
        const position = arrayReader.getPosition();
        const metadaArr = decoded.slice(
            position,
            position + metadataSize.byteLength
        );

        return uintToString(metadaArr);
    };

    private setMetadataXml = (metadataXml: string, metadata: Metadata) => {
        const parser = new DOMParser();
        const serializer = new XMLSerializer();
        const parsedMetadata = parser.parseFromString(metadataXml, "text/xml");
        if (metadata.queryName) {
            const items = parsedMetadata.getElementsByTagName("ItemPath");
            if (items && items.length) {
                for (let i = 0; i < items.length; i++) {
                    const item = items[i];
                    const content = item.innerHTML;
                    if (content.includes("Section1/")) {
                        const strArr = content.split("/");
                        strArr[1] = metadata.queryName;
                        const newContent = strArr.join("/");
                        item.innerHTML = newContent;
                        //append child
                    }
                }
            }
        }

        return serializer.serializeToString(parsedMetadata);
    };
}

function uintToString(uintArray: Uint8Array) {
    var encodedString = String.fromCharCode.apply(null, uintArray as any);
    // const decodedString = decodeURIComponent(escape(encodedString));
    return encodedString;
}
