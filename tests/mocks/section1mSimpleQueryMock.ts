export const simpleQueryMock = 
    `shared newQueryName = 
    let
    Source = Folder.Files("C:\\Users\\user1\\Desktop\\test")
in
    Source`;

export const section1mSimpleQueryMock = 
        `section Section1;

        shared Query1 =
    ${simpleQueryMock};`;

export const section1mBlankQueryMock = `section Section1;\r\n\r\nshared Query1 = let\r\n    Source = ""\r\nin\r\n    Source;`;

export const section1mNewQueryNameSimpleMock = 
        `section Section1;

        shared newQueryName =
    ${simpleQueryMock};`;

export const section1mNewQueryNameBlankMock = `section Section1;\r\n\r\nshared newQueryName = let\r\n    Source = ""\r\nin\r\n    Source;`;
