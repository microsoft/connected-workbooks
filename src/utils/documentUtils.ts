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

const createNewElements = (doc: Document, data: string[], parentElem: Element, element: string, properties: string[] | null, values: Map<number, string[]> | null) : void => {
    // remove previous child nodes of parent except first child
    while (parentElem.firstChild != parentElem.lastChild) {
                if (parentElem.lastChild) {
                    parentElem.removeChild(parentElem.lastChild);
                }            
            }
        const oldElem = parentElem.firstChild;
        if (!oldElem) {
            throw new Error("Error in template");
        }

        // create elements 
        for (var elemIndex = 0; elemIndex < data.length; elemIndex++) {
            const elem = doc.createElementNS(doc.documentElement.namespaceURI, element);
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
}

export default { createOrUpdateProperty, getDocPropsProperties, createNewElements };
