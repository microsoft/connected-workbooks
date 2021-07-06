"use strict";
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
Object.defineProperty(exports, "__esModule", { value: true });
exports.concatArrays = exports.getInt32Buffer = exports.ArrayReader = void 0;
class ArrayReader {
    constructor(array) {
        this._array = array;
        this._position = 0;
    }
    getInt32() {
        const retVal = new DataView(this._array, this._position, 4).getInt32(0, true);
        this._position += 4;
        return retVal;
    }
    getBytes(bytes) {
        const retVal = this._array.slice(this._position, (bytes ? bytes + this._position : bytes));
        this._position += retVal.byteLength;
        return new Uint8Array(retVal);
    }
    reset() {
        this._position = 0;
    }
}
exports.ArrayReader = ArrayReader;
function getInt32Buffer(val) {
    const packageSizeBuffer = new ArrayBuffer(4);
    new DataView(packageSizeBuffer).setInt32(0, val, true);
    return new Uint8Array(packageSizeBuffer);
}
exports.getInt32Buffer = getInt32Buffer;
function concatArrays(...args) {
    let size = 0;
    args.forEach(arr => size += arr.byteLength);
    const retVal = new Uint8Array(size);
    let position = 0;
    args.forEach(arr => {
        retVal.set(arr, position);
        position += arr.byteLength;
    });
    return retVal;
}
exports.concatArrays = concatArrays;
//# sourceMappingURL=arrayUtils.js.map