// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

/**
 * Cross-platform DOM utilities that work in both Node.js and browser environments
 */

let DOMParserImpl: typeof DOMParser;
let XMLSerializerImpl: typeof XMLSerializer;

// Check if we're in a browser environment
if (typeof window !== 'undefined' && window.DOMParser && window.XMLSerializer) {
    // Browser environment - use native implementations
    DOMParserImpl = window.DOMParser;
    XMLSerializerImpl = window.XMLSerializer;
} else {
    // Node.js environment - use @xmldom/xmldom
    try {
        // eslint-disable-next-line @typescript-eslint/no-var-requires
        const xmldom = require('@xmldom/xmldom');
        DOMParserImpl = xmldom.DOMParser;
        XMLSerializerImpl = xmldom.XMLSerializer;
    } catch (error) {
        throw new Error('DOM implementation not found. Please ensure @xmldom/xmldom is installed for Node.js environments.');
    }
}

export { DOMParserImpl as DOMParser, XMLSerializerImpl as XMLSerializer };