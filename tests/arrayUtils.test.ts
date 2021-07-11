import { getInt32Buffer, concatArrays, ArrayReader } from "../src/arrayUtils";
import * as base64 from "byte-base64";

describe("ArrayReader tests", () => {
    const buffer = base64.base64ToBytes("UHJhaXNlIFRoZSBTdW4h").buffer;
    const arrReader = new ArrayReader(buffer);

    test("getInt32 test", () => {
        const int32 = arrReader.getInt32();

        expect(int32).toEqual(1767993936);
        expect((arrReader as any)._position).toEqual(4);
    });

    test("getBytes test", () => {
        const bytes = arrReader.getBytes(4);

        expect(bytes).toEqual(new Uint8Array([115, 101, 32, 84]));
        expect((arrReader as any)._position).toEqual(8);
    });

    test("reset test", () => {
        arrReader.reset();

        expect((arrReader as any)._position).toEqual(0);
    });
});

test("getInt32Buffer test", () => {
    const size = 4;
    const val = 4;

    const packageSizeBuffer = new ArrayBuffer(size);
    new DataView(packageSizeBuffer).setInt32(0, val, true);
    const expected = new Uint8Array(packageSizeBuffer);

    const actual = getInt32Buffer(size);

    expect(actual).toStrictEqual(expected);
});

test("concatArrays test", () => {
    const uIntArr1 = new Uint8Array(4).fill(5);
    const uIntArr2 = new Uint8Array(2).fill(10);
    const expected = new Uint8Array(6);

    expected.set(uIntArr1, 0);
    expected.set(uIntArr2, 4);

    const actual = concatArrays(uIntArr1, uIntArr2);

    expect(actual).toStrictEqual(expected);
});
