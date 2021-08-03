export const generateMashupXMLTemplate = (base64: string): string =>
    `<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">${base64}</DataMashup>`;

export const generateSection1mString = (
    queryName: string,
    query: string
): string =>
    `section Section1;
    
    shared ${queryName} = 
    ${query};`;
