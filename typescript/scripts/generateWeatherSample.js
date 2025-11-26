const path = require("path");
const fs = require("fs/promises");
const { workbookManager } = require("../dist");

const mashup = `let
    Source = #table(
        {"City","TempC"},
        {
            {"Seattle", 18},
            {"London", 15},
            {"Sydney", 22}
        }
    )
in
    Source`;

const query = {
    queryMashup: mashup,
    queryName: "WeatherSample",
    refreshOnOpen: false,
};

const grid = {
    data: [
        ["City", "TempC"],
        ["Seattle", 0],
        ["London", 0],
        ["Sydney", 0],
    ],
    config: { promoteHeaders: true },
};

async function main() {
    const blob = await workbookManager.generateSingleQueryWorkbook(query, grid);
    const buffer = Buffer.from(await blob.arrayBuffer());
    const outputPath = path.resolve(__dirname, "..", "..", "WeatherSample.ts.xlsx");
    await fs.writeFile(outputPath, buffer);
    console.log(`TypeScript workbook generated: ${outputPath}`);
}

main().catch((error) => {
    console.error("Failed to generate TypeScript workbook", error);
    process.exitCode = 1;
});
