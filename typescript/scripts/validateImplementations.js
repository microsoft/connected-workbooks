/* eslint-disable @typescript-eslint/no-var-requires */
const path = require("path");
const { execSync } = require("child_process");

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

    const inspectScript = path.resolve(tsRoot, "scripts", "inspectMashup.js");
    run(`node "${inspectScript}" "${dotnetWorkbook}" "${tsWorkbook}"`, repoRoot);
}

main().catch((error) => {
    console.error(error.message || error);
    process.exit(1);
});
