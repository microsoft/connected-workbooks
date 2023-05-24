import { WorkbookManager } from './workbookManager';

if (process.argv.length < 3) {
    console.log('Usage: node cli.js <path to workbook>');
    process.exit(1);
}
var filepath = process.argv[2];

const workbookManager = new WorkbookManager();
workbookManager.getMQueryData(filepath).then((queries) => {
    let metadata = { "timestamp": new Date().toISOString(), "version": "0.1.3", hostname: require("os").hostname() };
    let fileQueries = { filepath, queries, metadata };
    console.log(JSON.stringify(fileQueries));
}).catch((err) => {
    console.error(err);
});
