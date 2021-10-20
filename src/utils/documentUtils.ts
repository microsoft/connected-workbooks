// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { docPropsCoreXmlPath, docPropsRootElement } from "../constants";

export const createOrUpdateProperty = (
    doc: Document,
    parent: Element,
    property: string,
    value: string | null | undefined
): void => {
    if (value === undefined) {
        return;
    }

    const elements = parent.getElementsByTagName(property);
    if (elements.length === 0) {
        const newElement = doc.createElement(property);
        newElement.nodeValue = value;
        parent.appendChild(newElement);
    } else if (elements.length > 1) {
        throw new Error(
            `Invalid DocProps core.xml, multiple ${property} elements`
        );
    } else {
        elements.item(0)!.textContent = value;
    }
};

export const getDocPropsProperties = async (
    zip: JSZip
): Promise<{ doc: Document; properties: Element }> => {
    const docPropsCoreXmlString = await zip
        .file(docPropsCoreXmlPath)
        ?.async("text");
    if (docPropsCoreXmlString === undefined) {
        throw new Error("DocProps core.xml was not found in template");
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(docPropsCoreXmlString, "text/xml");

    const properties = doc.getElementsByTagName(docPropsRootElement).item(0);
    if (properties === null) {
        throw new Error("Invalid DocProps core.xml");
    }

    return { doc, properties };
};

export default { createOrUpdateProperty, getDocPropsProperties };
