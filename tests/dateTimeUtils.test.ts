import { DataTypes } from "../src/types";
import { dateTimeUtils } from "../src/utils/";
import { shortTimeFormat } from "../src/utils/constants";

describe("DateTime Utils tests", () => {
    
    test("Detect DateTime format success", () => {
        expect(dateTimeUtils.detectDateTimeFormat("1/24/1996")).toEqual(DataTypes.shortDate);
        expect(dateTimeUtils.detectDateTimeFormat("Wednesday, January 24 1996")).toEqual(DataTypes.longDate);
        expect(dateTimeUtils.detectDateTimeFormat("1:24 PM")).toEqual(DataTypes.shortTime);
        expect(dateTimeUtils.detectDateTimeFormat("1:24:00 PM")).toEqual(DataTypes.longTime);
    });

    test("Detect DateTime format failure", () => {
        expect(dateTimeUtils.detectDateTimeFormat("1/24/199")).toBeUndefined();
        expect(dateTimeUtils.detectDateTimeFormat("Wednesday, January 24 199")).toBeUndefined();
        expect(dateTimeUtils.detectDateTimeFormat("2/30/1996")).toBeUndefined();
        expect(dateTimeUtils.detectDateTimeFormat("1:24 P")).toBeUndefined();
        expect(dateTimeUtils.detectDateTimeFormat("1:24:00 P")).toBeUndefined();
        expect(dateTimeUtils.detectDateTimeFormat("1/24")).toBeUndefined();
        expect(dateTimeUtils.detectDateTimeFormat("24/1/1996")).toBeUndefined();
        expect(dateTimeUtils.detectDateTimeFormat("1/32/1996")).toBeUndefined();
    });

    test ("Convert to Excel Date success", () => {
        expect(dateTimeUtils.convertToExcelDate("1/2/1900", DataTypes.shortDate)).toEqual(2);
        expect(dateTimeUtils.convertToExcelDate("Sunday, January 2 1900",DataTypes.longDate)).toEqual(2);
        expect(dateTimeUtils.convertToExcelDate("1:00 AM",  DataTypes.shortTime)).toBeCloseTo(1/24);
        expect(dateTimeUtils.convertToExcelDate("0:01:00 AM", DataTypes.longTime)).toBeCloseTo(1/(24*60));

    });

});