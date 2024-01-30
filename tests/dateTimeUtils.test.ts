import { dateTimeUtils } from "../src/utils/";

describe("DateTime Utils tests", () => {
    
    test("Detect DateTime format success", () => {
        expect(dateTimeUtils.detectDateTimeFormat("1/24/1996")).toEqual(dateTimeUtils.dateTimeFormatArr[0]);
        expect(dateTimeUtils.detectDateTimeFormat("Wednesday, January 24 1996")).toEqual(dateTimeUtils.dateTimeFormatArr[1]);
        expect(dateTimeUtils.detectDateTimeFormat("1:24 PM")).toEqual(dateTimeUtils.dateTimeFormatArr[2]);
        expect(dateTimeUtils.detectDateTimeFormat("1:24:00 PM")).toEqual(dateTimeUtils.dateTimeFormatArr[3]);
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
        expect(dateTimeUtils.convertToExcelDate("1/2/1900", dateTimeUtils.dateTimeFormatArr[0])).toEqual(2);
        expect(dateTimeUtils.convertToExcelDate("Sunday, January 2 1900", dateTimeUtils.dateTimeFormatArr[1])).toEqual(2);
        expect(dateTimeUtils.convertToExcelDate("Sunday, March 1 1900", dateTimeUtils.dateTimeFormatArr[1])).toEqual(61);
        expect(dateTimeUtils.convertToExcelDate("Sunday, February 29 1900", dateTimeUtils.dateTimeFormatArr[1])).toEqual(61);
        expect(dateTimeUtils.convertToExcelDate("1:00 AM", dateTimeUtils.dateTimeFormatArr[2])).toBeCloseTo(1/24);
        expect(dateTimeUtils.convertToExcelDate("0:01:00 AM", dateTimeUtils.dateTimeFormatArr[3])).toBeCloseTo(1/(24*60));
    });

});