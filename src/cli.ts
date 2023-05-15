import { WorkbookManager } from './workbookManager';

var filepath = '';
filepath = 'data/workbook-no-mquery.xlsx';
filepath = 'data/workbook.xlsx';
filepath = 'data/workbook-two-queries.xlsx';

const workbookManager = new WorkbookManager();
workbookManager.getMQueryData(filepath).then((queries) => {
    console.log(JSON.stringify(queries));
}).catch((err) => {
    console.log(JSON.stringify(err));
});