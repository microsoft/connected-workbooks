import { DataTypes } from "../src/types";
import { documentUtils } from "../src/utils";


describe("Document Utils tests", () => {
    test("ResolveType date not supported success", async () => {
        expect(documentUtils.resolveType(DataTypes.autodetect, "5-4-2023 00:00")).toEqual(DataTypes.string);
    });

    test("ResolveType string success", async () => {
        expect(documentUtils.resolveType(DataTypes.autodetect, "sTrIng")).toEqual(DataTypes.string);
    });
    
    test("ResolveType boolean success", async () => {
        expect(documentUtils.resolveType(DataTypes.autodetect, "true")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType(DataTypes.autodetect, "   true")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType(DataTypes.autodetect, "True")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType(DataTypes.autodetect, "false")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType(DataTypes.autodetect, "False")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType(DataTypes.autodetect, "   False")).toEqual(DataTypes.boolean); 
    });

    test("ResolveType number success", async () => {
        expect(documentUtils.resolveType(DataTypes.autodetect, "100000")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType(DataTypes.autodetect, "1000.00")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType(DataTypes.autodetect, "1000.50")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType(DataTypes.autodetect, "1000.50")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType(DataTypes.autodetect, "1.23450")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType(DataTypes.autodetect, "1.23.4.50")).toEqual(DataTypes.string);
    });

    test("ResolveType not autoDetect success", async () => {
        expect(documentUtils.resolveType(DataTypes.number, 1)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType(DataTypes.boolean, false)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType(DataTypes.string, "string")).toEqual(DataTypes.string);
    });
});