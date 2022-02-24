// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

export const generateMashupXMLTemplate = (base64: string): string =>
    `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${base64}</DataMashup>`;

export const generateSingleQueryMashup = (queryName: string, query: string): string =>
    `section Section1;
    
    shared ${queryName} = 
    ${query};`;

export const generateCustomXmlFilePath = (i: number): string => `customXml/item${i}.xml`;
