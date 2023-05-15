import { WorkbookManager } from './workbookManager';

if (process.argv.length < 3) {
    console.log('Usage: node cli.js <path to workbook>');
    process.exit(1);
}
var filepath = process.argv[2];

const workbookManager = new WorkbookManager();
workbookManager.getMQueryData(filepath).then((queries) => {
    console.log(JSON.stringify(queries));
}).catch((err) => {
    console.error(err);
});