export const simpleQueryMock = `
    let
    Source = Folder.Files("C:\\Users\\user1\\Desktop\\test")
in
    Source`;

export const section1mSimpleQueryMock = `section Section1;

        shared Query1 =
    ${simpleQueryMock};`;

export const section1mBlankQueryMock = `section Section1;\r\n\r\nshared Query1 = let\r\n    Source = ""\r\nin\r\n    Source;`;
