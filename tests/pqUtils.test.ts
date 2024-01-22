import { pqUtils } from "../src/utils/";

describe("Pq Utils tests", () => {
        test("tests that validation fails when non unique query names are given", () => {
            try {
                pqUtils.validateMultipleQueryNames([{ queryName: "  QuErY2 ", queryMashup: "" }, { queryName: "queRy2 ", queryMashup: "" }], "Query1");
                // If the above line doesn't throw an error, the test fails
                expect(true).toEqual(false);
            } catch (e) {
                expect(e.message).toEqual("Queries must have unique names");
            }
            try {
                pqUtils.validateMultipleQueryNames([{ queryName: "    qUeRy1  ", queryMashup: "" }, { queryName: "Query2", queryMashup: "" }], "  QuERy1 ");
                // If the above line doesn't throw an error, the test fails
                expect(true).toEqual(false);
            } catch (e) {
                expect(e.message).toEqual("Queries must have unique names");
            }
        });

        test("tests that validation succeeds when valid unique query names are given", () => {
            try {
                pqUtils.validateMultipleQueryNames([{ queryName: "Query 1", queryMashup: "" }, { queryName: "Query1", queryMashup: "" }], "Query2");
                expect(true).toEqual(true);
            } catch (e) {
                // If the above line throws an error, the test fails
                expect(true).toEqual(false);
            }
            try {
                pqUtils.validateMultipleQueryNames([{ queryName: "Query 1", queryMashup: "" }, { queryName: "Query1", queryMashup: "" }], "Query   1");
                expect(true).toEqual(true);
            } catch (e) {
                // If the above line throws an error, the test fails
                expect(true).toEqual(false);
            }
        });
        
        test("tests generated query name", () => {
           expect(pqUtils.generateUniqueQueryName(["connection only query-1", "connection only query-2", "connection only query-4"], "  Connection only query-3  ")).toEqual("Connection only query-5");
           expect(pqUtils.generateUniqueQueryName(["connection only query-1", "connection only query-2", "connection only query-3"], "connection only query -4")).toEqual("Connection only query-4");
           expect(pqUtils.generateUniqueQueryName(["Connection only query - 1", "connection only query-2", "connection only query-3"], "connection only query-4")).toEqual("Connection only query-1");
        });
        
});