// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import * as base64 from "byte-base64";
import {ArrayReader, concatArrays, getInt32Buffer} from "./arrayUtils";

export default class MashupHandler {
  async ReplaceSingleQuery(base64str:string, query:string) {
    let buffer = base64.base64ToBytes(base64str).buffer;
    let mashupArray = new ArrayReader(buffer);
    let startArray = mashupArray.getBytes(4);
    let packageSize = mashupArray.getInt32();
    let packageOPC = mashupArray.getBytes(packageSize);
    let endBuffer = mashupArray.getBytes();
    let newPackageBuffer = await this.editSingleQueryPackage(packageOPC, "Query1", query);
    let packageSizeBuffer = getInt32Buffer(newPackageBuffer.byteLength);
    let newMashup = concatArrays(startArray, packageSizeBuffer, newPackageBuffer, endBuffer);
    return base64.bytesToBase64(newMashup);
  }

  private async editSingleQueryPackage(packageOPC:ArrayBuffer, queryName:string, query: string) {
    let packageZip = await JSZip.loadAsync(packageOPC);
    let section1m = await packageZip.file("Formulas/Section1.m")?.async("text");
    if (section1m === undefined) {
      throw new Error("Formula section wasn't found in template");
    }
    let newSection1m = 
    `section Section1;
    
    shared ${queryName} = 
    ${query};`;

    packageZip.file("Formulas/Section1.m", newSection1m, { compression: "" });

    return await packageZip.generateAsync({ type: "uint8array" });
  }
}
