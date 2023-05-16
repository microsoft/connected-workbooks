import workbookTemplate from "../src/workbookTemplate";
import { GridParser }  from "../src/GridParser";


describe("Grid Parser tests", () => {
    const gridParser = new GridParser() as any;

    test("Connection XML attributes contain new query name", async () => {
        const RandomExcelDate = gridParser.convertToExcelDate("5/4/2023 00:00");
        expect(RandomExcelDate).toEqual(45050);
        const FirstExcelDate = gridParser.convertToExcelDate("1/1/1900 00:00");
        expect(FirstExcelDate).toEqual(1);
    })


  

});