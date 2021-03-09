export declare class ArrayReader {
    private _array;
    private _position;
    constructor(array: ArrayBuffer);
    getInt32(): number;
    getBytes(bytes?: number): Uint8Array;
    reset(): void;
}
export declare function getInt32Buffer(val: number): Uint8Array;
export declare function concatArrays(...args: Uint8Array[]): Uint8Array;
