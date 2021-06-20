// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import * as base64 from "byte-base64";
import {ArrayReader, concatArrays, getInt32Buffer} from "./arrayUtils";

export default class MashupHandler {
  async ReplaceSingleQuery(base64str:string, query:string): Promise<string> {
    const buffer = base64.base64ToBytes(base64str).buffer;
    const mashupArray = new ArrayReader(buffer);
    const startArray = mashupArray.getBytes(4);
    const packageSize = mashupArray.getInt32();
    const packageOPC = mashupArray.getBytes(packageSize);
    const endBuffer = mashupArray.getBytes();
    const newPackageBuffer = await this.editSingleQueryPackage(packageOPC, "Query1", query);
    const packageSizeBuffer = getInt32Buffer(newPackageBuffer.byteLength);
    const newMashup = concatArrays(startArray, packageSizeBuffer, newPackageBuffer, endBuffer);
    return base64.bytesToBase64(newMashup);
  }

  private async editSingleQueryPackage(packageOPC:ArrayBuffer, queryName:string, query: string) {
    const packageZip = await JSZip.loadAsync(packageOPC);
    const section1m = await packageZip.file("Formulas/Section1.m")?.async("text");
    if (section1m === undefined) {
      throw new Error("Formula section wasn't found in template");
    }
    const newSection1m = 
    `section Section1;
    
    shared ${queryName} = 
    ${query};`;

    packageZip.file("Formulas/Section1.m", newSection1m, { compression: "" });

    return await packageZip.generateAsync({ type: "uint8array" });
  }
}
