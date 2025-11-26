/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const { extractWorkbookDetails } = require("./workbookDetails");

async function main() {
    const targets = process.argv.slice(2);
    if (targets.length === 0) {
        console.error("Usage: node inspectMashup.js <workbook> [workbook...]");
        process.exit(1);
    }

    for (const target of targets) {
        const fullPath = path.resolve(target);
        const details = await extractWorkbookDetails(fullPath);
        const sectionPreview = details.sectionContent.split("\n").slice(0, 5).join("\n");
        const metadataPreview = details.metadataXml.slice(0, 400);
        console.log(`\n${fullPath}`);
        console.log(`  Query name : ${details.queryName}`);
        console.log(`  Metadata   : ${details.metadataBytesLength ?? details.metadataXml.length} bytes`);
        console.log("  Section1.m preview:\n" + sectionPreview.split("\n").map((line) => `    ${line}`).join("\n"));
        console.log("  Metadata preview:\n" + metadataPreview.split("\n").map((line) => `    ${line}`).join("\n"));
    }
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
