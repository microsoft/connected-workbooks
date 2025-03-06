// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

function getInt32Buffer(val: number): Uint8Array {
    const packageSizeBuffer = new ArrayBuffer(4);
    new DataView(packageSizeBuffer).setInt32(0, val, true);
    return new Uint8Array(packageSizeBuffer);
}

function concatArrays(...args: Uint8Array[]): Uint8Array {
    let size = 0;
    args.forEach((arr) => (size += arr.byteLength));
    const retVal = new Uint8Array(size);
    let position = 0;
    args.forEach((arr) => {
        retVal.set(arr, position);
        position += arr.byteLength;
    });

    return retVal;
}

export default {
    getInt32Buffer,
    concatArrays,
};
