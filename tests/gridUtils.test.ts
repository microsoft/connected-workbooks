import { arrayIsntMxNErr, defaults, promotedHeadersCannotBeUsedWithoutAdjustingColumnNamesErr } from "../src/utils/constants";
import gridUtils from "../src/utils/gridUtils";

const columnName = (i: number) => `${defaults.columnName} ${i}`;

describe("Grid Utils tests", () => {
    test.concurrent.each([
        ["null grid", null, { columnNames: [columnName(1)], rows: [[""]] }],
        ["null grid data", { data: null }, { columnNames: [columnName(1)], rows: [[""]] }],
        ["empty grid 1", { data: [] }, { columnNames: [columnName(1)], rows: [[""]] }],
        ["empty grid 2", { data: [[]] }, { columnNames: [columnName(1)], rows: [[""]] }],
        ["empty grid with empty rows", { data: [[], []] }, { columnNames: [columnName(1)], rows: [[""], [""]] }],
        [
            "happy path, no headers",
            {
                data: [
                    ["1", "2"],
                    ["3", "4"],
                ],
            },
            {
                columnNames: [columnName(1), columnName(2)],
                rows: [
                    ["1", "2"],
                    ["3", "4"],
                ],
            },
        ],
        [
            "type conversions, no headers",
            {
                data: [
                    [true, 3],
                    ["3", "4"],
                ],
            },
            {
                columnNames: ["Column 1", "Column 2"],
                rows: [
                    ["true", "3"],
                    ["3", "4"],
                ],
            },
        ],
        [
            "fill in empty rows",
            { data: [["1", "2"], [], ["3", "4"]] },
            {
                columnNames: [columnName(1), columnName(2)],
                rows: [
                    ["1", "2"],
                    ["", ""],
                    ["3", "4"],
                ],
            },
        ],
        [
            "promote headers with empty array",
            { data: [], config: { promoteHeaders: true } },
            {
                columnNames: [columnName(1)],
                rows: [[""]],
            },
        ],
        [
            "promote headers with empty row",
            { data: [[]], config: { promoteHeaders: true } },
            {
                columnNames: [columnName(1)],
                rows: [[""]],
            },
        ],
        [
            "promote headers, basic",
            {
                data: [
                    ["1", "2"],
                    ["3", "4"],
                ],
                config: { promoteHeaders: true },
            },
            {
                columnNames: ["1", "2"],
                rows: [["3", "4"]],
            },
        ],
        [
            "promote headers with empty array, without adjust column names",
            { data: [], config: { promoteHeaders: true, adjustColumnNames: false } },
            {
                columnNames: [columnName(1)],
                rows: [[""]],
            },
        ],
        [
            "promote headers, adjust column names, basic",
            {
                data: [
                    ["A", "A", "B"],
                    ["1", "2", "3"],
                ],
                config: { promoteHeaders: true },
            },
            {
                columnNames: ["A", "A (1)", "B"],
                rows: [["1", "2", "3"]],
            },
        ],
        [
            "promote headers, adjust column names, multiple",
            {
                data: [
                    ["A", "A", "B", "C", "B", "A"],
                    [1, 2, 3, 4, 5, 6],
                    [7, 8, 9, 10, 11, 12],
                ],
                config: { promoteHeaders: true },
            },
            {
                columnNames: ["A", "A (1)", "B", "C", "B (1)", "A (2)"],
                rows: [
                    ["1", "2", "3", "4", "5", "6"],
                    ["7", "8", "9", "10", "11", "12"],
                ],
            },
        ],
        [
            "promote headers, adjust column names, types",
            { data: [[true, true]], config: { promoteHeaders: true } },
            {
                columnNames: ["true", "true (1)"],
                rows: [["", ""]],
            },
        ],
    ])("%s:\n\t%j should be parsed to %j", (scenario, input, expected) => {
        expect(gridUtils.parseToTableData(input)).toEqual(expected);
    });

    // promote headers, without adjust column names, errors
    test.concurrent.each<(string | number | boolean)[][][]>([
        [[[]]],
        [[["A", "B", "A"]]],
        [[["A", 3, "B", 3]]],
        [[[true, "true"]]],
        [
            [
                ["אבג", "אבג"],
                ["1", "2"],
            ],
        ],
    ])(`parsing %j should throw "${promotedHeadersCannotBeUsedWithoutAdjustingColumnNamesErr}"`, (input) => {
        expect(() => gridUtils.parseToTableData({ data: input, config: { promoteHeaders: true, adjustColumnNames: false } })).toThrowError(
            promotedHeadersCannotBeUsedWithoutAdjustingColumnNamesErr
        );
    });

    // array isn't MxN
    test.concurrent.each<(string | number | boolean)[][][]>([
        [
            [
                ["A", "B", "A"],
                ["1", "2"],
            ],
        ],
        [[["אבג", "אבג"], ["1", "2"], ["3"]]],
    ])(`parsing %j should throw "${arrayIsntMxNErr}"`, (input) => {
        expect(() => gridUtils.parseToTableData({ data: input, config: { promoteHeaders: true, adjustColumnNames: false } })).toThrowError(arrayIsntMxNErr);
    });
});
