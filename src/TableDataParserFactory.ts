import GridParser from "./GridParser";
import { TableDataParser } from "./types";

export default class TableDataParserFactory {
 public static createParser(data: any): TableDataParser {
    if (data.Header !== undefined && data.GridData !== undefined) {
        return new GridParser();
    }
    throw new Error("Unsupported data type");
 }
}