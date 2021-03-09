export declare class QueryInfo {
    queryMashup: string;
    refreshOnOpen: boolean;
    constructor(queryMashup: string, refreshOnOpen: boolean);
}
export declare class WorkbookManager {
    private mashupHandler;
    generateSingleQueryWorkbook(query: QueryInfo, templateFile?: File): Promise<Blob>;
    private generateSingleQueryWorkbookFromZip;
    private setSingleQueryRefreshOnOpen;
    private setBase64;
    private getBase64;
}
