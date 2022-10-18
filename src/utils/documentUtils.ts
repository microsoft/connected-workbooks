// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import JSZip from "jszip";
import { docPropsCoreXmlPath, docPropsRootElement } from "../constants";

const createOrUpdateProperty = (doc: Document, parent: Element, property: string, value?: string | null): void => {
    if (value === undefined) {
        return;
    }

    const elements = parent.getElementsByTagName(property);

    if (elements?.length === 0) {
        const newElement = doc.createElement(property);
        newElement.textContent = value;
        parent.appendChild(newElement);
    } else if (elements.length > 1) {
        throw new Error(`Invalid DocProps core.xml, multiple ${property} elements`);
    } else if (elements?.length > 0) {
        elements[0]!.textContent = value;
        elements[0]!.innerHTML = value!;
    }
};

const getDocPropsProperties = async (zip: JSZip): Promise<{ doc: Document; properties: Element }> => {
    const docPropsCoreXmlString = await zip.file(docPropsCoreXmlPath)?.async("text");
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

const createNewElements = (doc: Document, data: string[], parentElem: string, element: string, properties: string[] | null, values: Map<number, string[]> | null) : void => {
    // remove previous child nodes of parent except first child
    const parentArr = [...doc.getElementsByTagName(parentElem)];
    parentArr.forEach((parent) => {
        while (parent.firstChild != parent.lastChild) {
                if (parent.lastChild) {
                    parent.removeChild(parent.lastChild);
                }            
            }
        const oldElem = parent.firstChild;
        if (!oldElem) {
            throw new Error("Error in template");
        }
        // duplicate number of elements necessary
        for (var elemId = 1; elemId < data.length; elemId++) {
            const newElem = oldElem.cloneNode(true);
            parent.appendChild(newElem);
        }
        const elemArr = [...doc.getElementsByTagName(element)];
        // update element properties
        for (var elemIndex = 0; elemIndex < elemArr.length; elemIndex++) {
            const elem = elemArr[elemIndex];
            if (properties && values) {
            for (var propertyIndex = 0; propertyIndex < properties.length; propertyIndex++) {
                const elemValues = values.get(elemIndex);
                if (!elemValues) {
                    throw new Error("Invalid ValuesMap.");
                }
                elem.setAttribute(properties[propertyIndex], elemValues[propertyIndex]);
            } 
        }          
        }
});
}

export default { createOrUpdateProperty, getDocPropsProperties, createNewElements };
