// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import { ArrayReader, concatArrays, getInt32Buffer } from "./arrayUtils";
import zipUtils from "./zipUtils";
export default class MashupHandler {
    async ReplaceSingleQuery(
        base64str: string,
        query: string
    ): Promise<string> {
        const buffer = base64.toByteArray(base64str).buffer;
        const mashupArray = new ArrayReader(buffer);
        const startArray = mashupArray.getBytes(4);
        const packageSize = mashupArray.getInt32();
        const packageOPC = mashupArray.getBytes(packageSize);
        const endBuffer = mashupArray.getBytes();
        const newPackageBuffer = await this.editSingleQueryPackage(
            packageOPC,
            "Query1",
            query
        );
        const packageSizeBuffer = getInt32Buffer(newPackageBuffer.byteLength);
        const newMashup = concatArrays(
            startArray,
            packageSizeBuffer,
            newPackageBuffer,
            endBuffer
        );
        return base64.fromByteArray(newMashup);
    }

    private async editSingleQueryPackage(
        packageOPC: ArrayBuffer,
        queryName: string,
        query: string
    ) {
        const packageZip = await zipUtils.loadAsync(packageOPC);

        zipUtils.chackAndgetSection1m(packageZip);
        zipUtils.setSection1m(queryName, query, packageZip);

        return await packageZip.generateAsync({ type: "uint8array" });
    }
}
