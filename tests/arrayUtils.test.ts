// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { describe, test, expect } from '@jest/globals';
import { arrayUtils } from "../src/utils/";


test("getInt32Buffer test", () => {
    const size = 4;
    const val = 4;

    const packageSizeBuffer = new ArrayBuffer(size);
    new DataView(packageSizeBuffer).setInt32(0, val, true);
    const expected = new Uint8Array(packageSizeBuffer);

    const actual = arrayUtils.getInt32Buffer(size);

    expect(actual).toStrictEqual(expected);
});

test("concatArrays test", () => {
    const uIntArr1 = new Uint8Array(4).fill(5);
    const uIntArr2 = new Uint8Array(2).fill(10);
    const expected = new Uint8Array(6);

    expected.set(uIntArr1, 0);
    expected.set(uIntArr2, 4);

    const actual = arrayUtils.concatArrays(uIntArr1, uIntArr2);

    expect(actual).toStrictEqual(expected);
});
