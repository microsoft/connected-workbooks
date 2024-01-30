// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import { ConnectionOnlyQueryInfo, QueryInfo } from "./types";
import { missingQueryNameErr } from "./utils/constants";

export const generateMashupXMLTemplate = (base64: string): string =>
    `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${base64}</DataMashup>`;

export const generateSingleQueryMashup = (queryName: string, query: string): string =>
    `section Section1;
    
    shared #"${queryName}" = 
    ${query};`;

export const generateMultipleQueryMashup = (loadedQuery: QueryInfo, queries: ConnectionOnlyQueryInfo[]): string => {
    if (!loadedQuery.queryName) {
        throw new Error(missingQueryNameErr);
    }

    let mashup: string = generateSingleQueryMashup(loadedQuery.queryName, loadedQuery.queryMashup);
    queries.forEach((query: ConnectionOnlyQueryInfo) => {
        const queryName = query.queryName;
        if (!queryName) {
            throw new Error(missingQueryNameErr);
        }

        mashup += `
        
        shared #"${queryName}" = 
        ${query.queryMashup};`;
    });
    
    return mashup;
}

export const generateCustomXmlFilePath = (i: number): string => `customXml/item${i}.xml`;
