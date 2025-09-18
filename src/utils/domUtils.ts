// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Cross-platform DOM utilities that work in both Node.js and browser environments
 */

// TypeScript types for DOMParser and XMLSerializer constructors
type DOMParserConstructor = new () => DOMParser;
type XMLSerializerConstructor = new () => XMLSerializer;

let domParserCache: DOMParserConstructor | undefined;
let xmlSerializerCache: XMLSerializerConstructor | undefined;

/**
 * Gets a DOMParser implementation that works in both browser and Node.js environments
 * In browsers, uses the native DOMParser
 * In Node.js, requires the optional @xmldom/xmldom dependency
 */
export function getDOMParser(): DOMParserConstructor {
    if (domParserCache) {
        return domParserCache;
    }

    // Browser environment - use native implementation
    if (typeof window !== 'undefined' && window.DOMParser) {
        domParserCache = window.DOMParser;
        return window.DOMParser;
    }

    // Node.js environment - try to load @xmldom/xmldom
    try {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const xmldom = require('@xmldom/xmldom');
        domParserCache = xmldom.DOMParser;
        return xmldom.DOMParser;
    } catch (error) {
        throw new Error(
            'DOM implementation not available in Node.js environment. ' +
            'Please install the optional dependency: npm install @xmldom/xmldom'
        );
    }
}

/**
 * Gets an XMLSerializer implementation that works in both browser and Node.js environments
 * In browsers, uses the native XMLSerializer
 * In Node.js, requires the optional @xmldom/xmldom dependency
 */
export function getXMLSerializer(): XMLSerializerConstructor {
    if (xmlSerializerCache) {
        return xmlSerializerCache;
    }

    // Browser environment - use native implementation
    if (typeof window !== 'undefined' && window.XMLSerializer) {
        xmlSerializerCache = window.XMLSerializer;
        return window.XMLSerializer;
    }

    // Node.js environment - try to load @xmldom/xmldom
    try {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const xmldom = require('@xmldom/xmldom');
        xmlSerializerCache = xmldom.XMLSerializer;
        return xmldom.XMLSerializer;
    } catch (error) {
        throw new Error(
            'DOM implementation not available in Node.js environment. ' +
            'Please install the optional dependency: npm install @xmldom/xmldom'
        );
    }
}

// For backward compatibility, export the classes directly
// These will throw helpful error messages if @xmldom/xmldom is not available in Node.js
export const DOMParser: DOMParserConstructor = getDOMParser();
export const XMLSerializer: XMLSerializerConstructor = getXMLSerializer();