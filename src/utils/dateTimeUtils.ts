import { DateTimeFormat } from "../types";
import { daysOfWeek, defaultDate, invalidDateTimeErr, longDateFormat, longDateReg, longTimeFormat, longTimeReg, milliSecPerDay, months, monthsbeforeLeap, numberOfDaysTillExcelBeginYear,  shortDateId,  shortDateReg,  shortTimeFormat, shortTimeReg } from "./constants";

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


// Validation functions for supported datetime formats
const longDateValidation = (data: string): boolean => {
    const dateParts: string[] = data.split(', ');
    if (dateParts.length != 2) {
        return false;
    }   

    const parsedDayofWeek: Number = daysOfWeek.indexOf(dateParts[0]);
    let parsedData: number[] = dateParts[1].split(' ').map((x) => Number(x) ? Number(x) : months.indexOf(x));
    const date: Date = new Date(parsedData![2], parsedData![0], parsedData![1]);

    return (date.getMonth() == parsedData![0] && date.getDate() == parsedData![1] && date.getFullYear() == parsedData![2] && date.getDay() == parsedDayofWeek);
}

const shortDateValidation = (data: string): boolean => {
    const parsedData = data.split("/").map((x) => Number(x));
    const date = new Date(data);
    if (parsedData.length != 3) {
        return false;
    }

    return (date.getMonth() + 1 == parsedData[0] && date.getDate() == parsedData[1] && date.getFullYear() == parsedData[2]);
}

// Supported datetime formats
export const dateTimeFormatArr: DateTimeFormat[] = 
    [{regex: shortDateReg, validateDate: shortDateValidation, formatId: shortDateId, isDate: true}, // M/d/yyyy
     {regex: longDateReg, validateDate: longDateValidation, formatCode: longDateFormat, isDate: true}, // dddd, mmmm dd, yyyy
     {regex: shortTimeReg, validateDate: (data: string) => true, formatCode: shortTimeFormat, isDate: false}, // h:mm AM/PM
     {regex: longTimeReg, validateDate: (data: string) => true, formatCode: longTimeFormat, isDate: false}, // h:mm:ss AM/PM
    ];

export default {
    convertToExcelDate,
    detectDateTimeFormat,
    dateTimeFormatArr
};