// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { JSZip } from "./zipUtils";
import { generateSection1mString } from "../generators";
import { section1mPath } from "../constants";

const setSection1m = (queryName: string, query: string, zip: JSZip): void => {
    const newSection1m = generateSection1mString(queryName, query);

    zip.file(section1mPath, newSection1m, {
        compression: "",
    });
};

const getSection1m = async (zip: JSZip): Promise<string> => {
    const section1m = zip.file(section1mPath)?.async("text");
    if (!section1m) {
        throw new Error("Formula section wasn't found in template");
    }

    return section1m;
};

export default {
    setSection1m,
    getSection1m,
};
