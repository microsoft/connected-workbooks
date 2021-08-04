// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import * as base64 from "base64-js";
import { arrayUtils, zipUtils, pqUtils } from "./utils/";

export default class MashupHandler {
    async ReplaceSingleQuery(
        base64str: string,
        query: string
    ): Promise<string> {
        const buffer = base64.toByteArray(base64str).buffer;
        const mashupArray = new arrayUtils.ArrayReader(buffer);
        const startArray = mashupArray.getBytes(4);
        const packageSize = mashupArray.getInt32();
        const packageOPC = mashupArray.getBytes(packageSize);
        const endBuffer = mashupArray.getBytes();
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

    private async editSingleQueryPackage(
        packageOPC: ArrayBuffer,
        queryName: string,
        query: string
    ) {
        const packageZip = await zipUtils.loadAsync(packageOPC);
        pqUtils.getSection1m(packageZip);
        pqUtils.setSection1m(queryName, query, packageZip);

        return await packageZip.generateAsync({ type: "uint8array" });
    }
}
