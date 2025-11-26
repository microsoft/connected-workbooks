/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const { extractWorkbookDetails } = require("./workbookDetails");

function usage() {
    console.error("Usage:");
    console.error("  node inspectMashup.js --pretty <workbook>");
    console.error("  node inspectMashup.js <workbookA> <workbookB>");
}

function printPretty(details) {
    const sectionPreview = details.sectionContent.split("\n").slice(0, 5).join("\n");
    const metadataPreview = details.metadataXml.slice(0, 400);
    console.log(`\n${details.workbookPath}`);
    console.log(`  Query name : ${details.queryName}`);
    console.log(`  Metadata   : ${details.metadataBytesLength ?? details.metadataXml.length} bytes`);
    console.log("  Section1.m preview:\n" + sectionPreview.split("\n").map((line) => `    ${line}`).join("\n"));
    console.log("  Metadata preview:\n" + metadataPreview.split("\n").map((line) => `    ${line}`).join("\n"));
}

function compare(detailsA, detailsB) {
    const checks = [
        ["Query name", detailsA.queryName, detailsB.queryName],
        ["Connection name", detailsA.connection.name, detailsB.connection.name],
        ["Connection description", detailsA.connection.description, detailsB.connection.description],
        ["Connection location", detailsA.connection.location, detailsB.connection.location],
        ["Connection command", detailsA.connection.command, detailsB.connection.command],
        ["Refresh flag", detailsA.connection.refreshOnLoad, detailsB.connection.refreshOnLoad],
    ];

    const mismatches = [];
    for (const [label, left, right] of checks) {
        if (left !== right) {
            mismatches.push(`${label} mismatch:\n    ${detailsA.workbookPath}: ${left}\n    ${detailsB.workbookPath}: ${right}`);
        }
    }

    const section1Filter = (paths) => paths.filter((entry) => entry.includes("Section1/"));
    const leftPaths = section1Filter(detailsA.metadataItemPaths);
    const rightPaths = section1Filter(detailsB.metadataItemPaths);
    if (JSON.stringify(leftPaths) !== JSON.stringify(rightPaths)) {
        mismatches.push("Metadata ItemPath entries differ");
    }

    if (!detailsA.sharedStrings.includes(detailsA.queryName)) {
        mismatches.push(`${path.basename(detailsA.workbookPath)} sharedStrings missing query name`);
    }

    if (!detailsB.sharedStrings.includes(detailsB.queryName)) {
        mismatches.push(`${path.basename(detailsB.workbookPath)} sharedStrings missing query name`);
    }

    return mismatches;
}

async function prettyMode(workbook) {
    const details = await extractWorkbookDetails(path.resolve(workbook));
    printPretty(details);
}

async function comparisonMode(left, right) {
    const leftDetails = await extractWorkbookDetails(path.resolve(left));
    const rightDetails = await extractWorkbookDetails(path.resolve(right));
    const mismatches = compare(leftDetails, rightDetails);
    if (mismatches.length === 0) {
        console.log("\n✅ Workbooks match for inspected fields.");
        return;
    }

    console.error("\n❌ Workbooks differ:");
    mismatches.forEach((item) => console.error(`  - ${item}`));
    process.exit(1);
}

async function main() {
    const args = process.argv.slice(2);
    if (args.length === 0) {
        usage();
        process.exit(1);
    }

    if (args[0] === "--pretty") {
        if (args.length !== 2) {
            usage();
            process.exit(1);
        }

        await prettyMode(args[1]);
        return;
    }

    if (args.length === 2) {
        await comparisonMode(args[0], args[1]);
        return;
    }

    usage();
    process.exit(1);
}

main().catch((error) => {
    console.error(error);
    process.exit(1);
});
