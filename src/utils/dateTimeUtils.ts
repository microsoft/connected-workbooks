import { DataTypes } from "../types";
import { dateTimeRegexes, daysOfWeek, defaultDate, invalidDateTimeErr, longDateFormat, longTimeFormat, milliSecPerDay, months, monthsbeforeLeap, numberOfDaysTillExcelBeginYear,  shortTimeFormat } from "./constants";


const convertToExcelDate = (dateTime: string, dataType: DataTypes) => {
    let dateTimeStr: string = dateTime;
    if (dataType == DataTypes.shortTime || dataType == DataTypes.longTime) {
        // Excel assumes that the time is in the 31/12/1899 if none is specified
        dateTimeStr = defaultDate + dateTimeStr;
    }

    const localDate: Date = new Date(dateTimeStr);
    if (isNaN(localDate.getTime())) {
        throw new Error(invalidDateTimeErr);
    }

    const globalDate = new Date(Date.UTC(localDate.getFullYear(), localDate.getMonth(), 
        localDate.getDate(), localDate.getHours(), localDate.getMinutes(), localDate.getSeconds())).getTime();
    // Excel incorrectly assumes that the year 1900 is a leap year. This is a workaround for that
    if ((localDate.getFullYear() == 1900 && localDate.getMonth() <= monthsbeforeLeap) || localDate.getFullYear() < 1900) {
        return ((globalDate + numberOfDaysTillExcelBeginYear*milliSecPerDay) / (milliSecPerDay)) - 1;
    }

    return (globalDate + numberOfDaysTillExcelBeginYear*milliSecPerDay) / (milliSecPerDay);
};

const getFormatCode = (format: DataTypes): string => {
    switch (format) {
        case DataTypes.longDate:
            return longDateFormat;
        case DataTypes.longTime:
            return longTimeFormat;
        case DataTypes.shortTime:
            return shortTimeFormat;
        default:
           throw new Error (invalidDateTimeErr);
    }
};

export const detectDateTimeFormat = (dateTime: string): DataTypes|undefined => { 
    for (const [regex, dataType] of dateTimeRegexes) {
        if (regex.test(dateTime)) {
            return validateDate(dateTime, dataType) ? dataType : undefined;
        }
    }
};

const validateDate = (data: string, dataType: DataTypes): boolean => {
    const date: Date = new Date(data);
    switch(dataType) {
        case DataTypes.shortDate:  
            const [parsedMonth, parsedDay, parsedYear] = data.split("/").map((x) => Number(x));
            
            return (date.getDate() == parsedDay && date.getMonth() + 1 == parsedMonth && date.getFullYear() == parsedYear); 

        case DataTypes.longDate:
            const dateParts: string[] = data.split(', ');
            if (dateParts.length != 2) {
                return false;
            }   
            // Extract day of the week
            const dayOfWeek: string = dateParts[0];
            // Extract day, month, and year
            const [parsedLongMonth, parsedLongDay, parsedLongYear] = dateParts[1].split(' ');
            
            return (date.getDate() == Number(parsedLongDay) && date.getMonth() == months.indexOf(parsedLongMonth) && date.getFullYear() == Number(parsedLongYear) && date.getDay()== daysOfWeek.indexOf(dayOfWeek));

        case DataTypes.shortTime:
            return true;
        case DataTypes.longTime:
            return true;    
        default:
            return false;
    }
}; 

export default {
    convertToExcelDate,
    getFormatCode,
    detectDateTimeFormat,
};