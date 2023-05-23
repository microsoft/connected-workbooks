// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { QueryInfo } from "./types";

export const generateMashupXMLTemplate = (base64: string): string =>
    `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${base64}</DataMashup>`;

export const generateSingleQueryMashup = (queryName: string, query: string): string =>
    `section Section1;
    
    shared ${queryName} = 
    ${query};`;

export const generateMultipleQueryMashup = (queryName: string, queryMashup: string, connectionOnlyQueries: QueryInfo[]): string => {
    let section1m: string =  generateSingleQueryMashup(queryName, queryMashup);  
    connectionOnlyQueries.forEach(query => {
        if (query.queryName === undefined) {
            throw new Error("Query name is undefined");
        }
        section1m += 
        `shared ${query.queryName!} = 
        ${query.queryMashup};`
    })
    return section1m;
}
    
export const generateCustomXmlFilePath = (i: number): string => `customXml/item${i}.xml`;
