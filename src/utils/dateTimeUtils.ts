import { DataTypes } from "../types";
import { defaultDate, invalidDateTimeErr, longDateFormat, longDateReg, longDateValidation, longTimeFormat, longTimeReg, milliSecPerDay, monthsbeforeLeap, numberOfDaysTillExcelBeginYear,  shortDateId,  shortDateReg, shortDateValidation,  shortTimeFormat, shortTimeReg } from "./constants";

export interface DateTimeFormat {
    regex: RegExp;
    validateDate: (data: string) => boolean;
    isDate: boolean;
    formatCode?: string;
    formatId?: number;
}

// Supported datetime formats
export const dateTimeFormatArr: DateTimeFormat[] = 
    [{regex: shortDateReg, validateDate: shortDateValidation, formatId: shortDateId, isDate: true}, // M/d/yyyy
     {regex: longDateReg, validateDate: longDateValidation, formatCode: longDateFormat, isDate: true}, // dddd, mmmm dd, yyyy
     {regex: shortTimeReg, validateDate: (data: string) => true, formatCode: shortTimeFormat, isDate: false}, // h:mm AM/PM
     {regex: longTimeReg, validateDate: (data: string) => true, formatCode: longTimeFormat, isDate: false}, // h:mm:ss AM/PM
    ];

const convertToExcelDate = (data: string, dateTime: DateTimeFormat) => {
    let dataStr: string = data;
    if (!dateTime.isDate) {
        // Excel assumes that the time is in the 31/12/1899 if none is specified
        dataStr = defaultDate + dataStr;
    }

    const localDate: Date = new Date(dataStr);
    if (isNaN(localDate.getTime())) {
        throw new Error(invalidDateTimeErr);
    }

    const globalDate = new Date(Date.UTC(localDate.getFullYear(), localDate.getMonth(), 
        localDate.getDate(), localDate.getHours(), localDate.getMinutes(), localDate.getSeconds())).getTime();
    // Excel incorrectly assumes that the year 1900 is a leap year. This is a workaround for that
    if ((localDate.getFullYear() == 1900 && localDate.getMonth() < monthsbeforeLeap) || localDate.getFullYear() < 1900) {
        return ((globalDate + numberOfDaysTillExcelBeginYear*milliSecPerDay) / (milliSecPerDay)) - 1;
    }

    return (globalDate + numberOfDaysTillExcelBeginYear*milliSecPerDay) / (milliSecPerDay);
};

export const detectDateTimeFormat = (data: string): DateTimeFormat|undefined => { 
    let dateTimeFormat: DateTimeFormat|undefined = undefined;
    Object.values(dateTimeFormatArr).forEach((format) => {
        if (format.regex.test(data) && format.validateDate(data)) {
            dateTimeFormat = format;
            
            return;
        }
    });
    
    return dateTimeFormat;
};

export const getFormatCode = (dataType: DataTypes): string|undefined => {
    if (dateTimeFormatArr[dataType] && dateTimeFormatArr[dataType]!.formatCode) {
        return dateTimeFormatArr[dataType]!.formatCode;
    }

    return undefined;
}

export default {
    convertToExcelDate,
    detectDateTimeFormat,
    getFormatCode,
    dateTimeFormatArr
};