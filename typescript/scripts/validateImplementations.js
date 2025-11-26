/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const { execSync } = require("child_process");
const { extractWorkbookDetails } = require("./workbookDetails");

const repoRoot = path.resolve(__dirname, "..", "..");
const tsRoot = path.resolve(__dirname, "..");

function run(command, cwd) {
    execSync(command, { cwd, stdio: "inherit" });
}

async function ensureSamples() {
    run("dotnet run --project dotnet/sample/ConnectedWorkbooks.Sample.csproj", repoRoot);
    run("npm run build", tsRoot);
    run("node scripts/generateWeatherSample.js", tsRoot);
}

function compare(detailsA, detailsB) {
    const mismatches = [];
    const checks = [
        ["Query name", detailsA.queryName, detailsB.queryName],
        ["Connection name", detailsA.connection.name, detailsB.connection.name],
        ["Connection description", detailsA.connection.description, detailsB.connection.description],
        ["Connection location", detailsA.connection.location, detailsB.connection.location],
        ["Connection command", detailsA.connection.command, detailsB.connection.command],
        ["Refresh flag", detailsA.connection.refreshOnLoad, detailsB.connection.refreshOnLoad],
    ];

    for (const [label, left, right] of checks) {
        if (left !== right) {
            mismatches.push(`${label} mismatch: '${left}' vs '${right}'`);
        }
    }

    const leftPaths = detailsA.metadataItemPaths.filter((entry) => entry.includes("Section1/"));
    const rightPaths = detailsB.metadataItemPaths.filter((entry) => entry.includes("Section1/"));
    if (JSON.stringify(leftPaths) !== JSON.stringify(rightPaths)) {
        mismatches.push("Metadata ItemPath entries differ");
    }

    if (!detailsA.sharedStrings.includes(detailsA.queryName)) {
        mismatches.push(".NET sharedStrings missing query name");
    }

    if (!detailsB.sharedStrings.includes(detailsB.queryName)) {
        mismatches.push("TypeScript sharedStrings missing query name");
    }

    if (mismatches.length > 0) {
        const error = mismatches.join("\n  - ");
        throw new Error(`Workbooks are not aligned:\n  - ${error}`);
    }
}

async function main() {
    const args = process.argv.slice(2);
    let dotnetWorkbook;
    let tsWorkbook;

    if (args.length === 2) {
        [dotnetWorkbook, tsWorkbook] = args.map((arg) => path.resolve(arg));
    } else if (args.length === 0) {
        await ensureSamples();
        dotnetWorkbook = path.resolve(repoRoot, "dotnet", "WeatherSample.xlsx");
        tsWorkbook = path.resolve(repoRoot, "WeatherSample.ts.xlsx");
    } else {
        console.error("Usage: node scripts/validateImplementations.js [dotnetWorkbook tsWorkbook]");
        process.exit(1);
    }

    const dotnetDetails = await extractWorkbookDetails(dotnetWorkbook);
    const tsDetails = await extractWorkbookDetails(tsWorkbook);
    compare(dotnetDetails, tsDetails);
    console.log("\n✅ Validation succeeded. The .NET and TypeScript outputs match for the inspected fields.");
}

main().catch((error) => {
    console.error(error.message || error);
    process.exit(1);
});
