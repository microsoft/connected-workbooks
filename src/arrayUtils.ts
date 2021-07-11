// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export class ArrayReader {
    private _array: ArrayBuffer;
    private _position: number;

    constructor(array: ArrayBuffer) {
        this._array = array;
        this._position = 0;
    }

    public getInt32() {
        const retVal = new DataView(this._array, this._position, 4).getInt32(
            0,
            true
        );
        this._position += 4;

        return retVal;
    }

    getBytes(bytes?: number): Uint8Array {
        const retVal = this._array.slice(
            this._position,
            bytes ? bytes! + this._position : bytes
        );
        this._position += retVal.byteLength;
        return new Uint8Array(retVal);
    }

    reset() {
        this._position = 0;
    }
}

export function getInt32Buffer(val: number) {
    const packageSizeBuffer = new ArrayBuffer(4);
    new DataView(packageSizeBuffer).setInt32(0, val, true);
    return new Uint8Array(packageSizeBuffer);
}

export function concatArrays(...args: Uint8Array[]) {
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
