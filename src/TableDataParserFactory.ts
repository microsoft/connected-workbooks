import GridParser from "./GridParser";
import { Grid, TableDataParser } from "./types";

export default class TableDataParserFactory {
 public static createParser(data: any): TableDataParser {
    if (data as Grid) {
        return new GridParser();
    }
    
    throw new Error("Unsupported data type");
 }
}