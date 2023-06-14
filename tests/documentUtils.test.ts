import { DataTypes } from "../src/types";
import { documentUtils } from "../src/utils";


describe("Document Utils tests", () => {
    test("ResolveType date not supported success", async () => {
        expect(documentUtils.resolveType("5-4-2023 00:00")).toEqual(DataTypes.string);
    });

    test("ResolveType string success", async () => {
        expect(documentUtils.resolveType("sTrIng")).toEqual(DataTypes.string);
    });
    
    test("ResolveType boolean success", async () => {
        expect(documentUtils.resolveType("true")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType( "   true")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("True")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("false")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("False")).toEqual(DataTypes.boolean);
        expect(documentUtils.resolveType("   False")).toEqual(DataTypes.boolean); 
    });

    test("ResolveType number success", async () => {
        expect(documentUtils.resolveType("100000")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.00")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1000.50")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23450")).toEqual(DataTypes.number);
        expect(documentUtils.resolveType("1.23.4.50")).toEqual(DataTypes.string);
    });

});