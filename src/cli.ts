import { WorkbookManager } from './workbookManager';

var filepath = '';
filepath = 'data/workbook-no-mquery.xlsx';
filepath = 'data/workbook.xlsx';

const workbookManager = new WorkbookManager();
workbookManager.getQueryInfo(filepath).then((queryInfo) => {
    console.log(JSON.stringify({ "query": queryInfo }));
}).catch((err) => {
    console.log(JSON.stringify(err));
});