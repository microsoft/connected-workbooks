import { DataTypes } from "../src/types";
import { documentUtils } from "../src/utils";


describe("Document Utils tests", () => {
    test("ResolveType date not supported success", async () => {
        expect(documentUtils.resolveType("5-4-2023 00:00", false)).toEqual(DataTypes.string);
    });

    test("ResolveType string success", async () => {
        expect(documentUtils.resolveType("sTrIng", false)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("True", false)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("False", false)).toEqual(DataTypes.string);
    });
    
    test("ResolveType boolean success", async () => {
        expect(documentUtils.resolveType("true", false)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType( "   true", false)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("false", false)).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("   false", false)).toEqual(DataTypes.boolean); 
    });

    test("ResolveType number success", async () => {
        expect(documentUtils.resolveType("100000", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.00", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23450", false)).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23.4.50", false)).toEqual(DataTypes.string);
    });

    test("ResolveType header row success", async () => {
        expect(documentUtils.resolveType("100000", true)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("true", true)).toEqual(DataTypes.string);
        expect(documentUtils.resolveType("string", true)).toEqual(DataTypes.string);
    });
});