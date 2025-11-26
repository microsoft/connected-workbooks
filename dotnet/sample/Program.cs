using Microsoft.ConnectedWorkbooks;
using Microsoft.ConnectedWorkbooks.Models;

var manager = new WorkbookManager();

var mashup = """
let
    Source = #table(
        {"City","TempC"},
        {
            {"Seattle", 18},
            {"London", 15},
            {"Sydney", 22}
        }
    )
in
    Source
""";

var query = new QueryInfo(
    queryMashup: mashup,
    queryName: "WeatherSample",
    refreshOnOpen: false);

var grid = new Grid(new[]
{
    new object?[] { "City", "TempC" },
    new object?[] { "Seattle", 0 },
    new object?[] { "London", 0 },
    new object?[] { "Sydney", 0 }
}, new GridConfig { PromoteHeaders = true });

var bytes = manager.GenerateSingleQueryWorkbook(query, grid);
var repoRoot = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, "..", "..", "..", ".."));
var outputPath = Path.Combine(repoRoot, "WeatherSample.xlsx");
await File.WriteAllBytesAsync(outputPath, bytes);

Console.WriteLine($"Workbook generated: {outputPath}");
