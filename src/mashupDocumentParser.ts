// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import JSZip from "jszip";
import { section1mPath } from "./constants";
import { arrayUtils } from "./utils";
import { generateSingleQueryMashup } from "./generators";

export default class MashupHandler {
    async ReplaceSingleQuery(
        base64Str: string,
        query: string
    ): Promise<string> {
        const { packageOPC, startArray, endBuffer } =
            this.getPackageComponents(base64Str);

        const newPackageBuffer = await this.editSingleQueryPackage(
            packageOPC,
            "Query1",
            query
        );
        const packageSizeBuffer = arrayUtils.getInt32Buffer(
            newPackageBuffer.byteLength
        );
        const newMashup = arrayUtils.concatArrays(
            startArray,
            packageSizeBuffer,
            newPackageBuffer,
            endBuffer
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
}
