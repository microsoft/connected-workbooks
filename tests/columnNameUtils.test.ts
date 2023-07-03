import columnNameUtils from "../src/utils/columnNameUtils";

describe("Grid Utils tests", () => {
    test("tests Get next column Name", () => {
        const columnNames = ["Column", "Column (1)", "Column (2)", "Column (3)", "Banana"];
        expect(columnNameUtils.getNextAvailableColumnName(columnNames, "Column")).toEqual("Column (4)");
        expect(columnNameUtils.getNextAvailableColumnName(columnNames, "Banana")).toEqual("Banana (1)");
        expect(columnNameUtils.getNextAvailableColumnName(columnNames, "unexists")).toEqual("unexists");
    });

    test("tests Get adjusted column Name", () => {
        expect(columnNameUtils.getAdjustedColumnNames(["Column", 2, false, 2, "Banana"])).toEqual(["Column", "2", "false", "2 (1)", "Banana"]);
        expect(columnNameUtils.getAdjustedColumnNames(["Column", 2, "", 2, "Banana"])).toEqual(["Column", "2", "Column (1)", "2 (1)", "Banana"]);
        expect(columnNameUtils.getAdjustedColumnNames(["Column", "Column", "Column (2)", "Column (3)"])).toEqual([
            "Column",
            "Column (1)",
            "Column (2)",
            "Column (3)",
        ]);
    });

    test("tests Get raw column Name success", () => {
        expect(columnNameUtils.getRawColumnNames(["Column", "Column1", "Column2", 2])).toEqual(["Column", "Column1", "Column2", "2"]);
    });

    test("tests Get raw column Name empty name", () => {
        try {
            columnNameUtils.getRawColumnNames(["Column", 2, false, 2, "Banana"]);
        } catch (error) {
            expect(error).toBeTruthy();
        }
    });

    test("tests Get raw column Name conflict name", () => {
        try {
            columnNameUtils.getRawColumnNames(["Column", 2, "", 2, "Banana"]);
        } catch (error) {
            expect(error).toBeTruthy();
        }
    });
});
