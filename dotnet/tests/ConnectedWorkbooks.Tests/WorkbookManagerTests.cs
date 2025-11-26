// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Microsoft.ConnectedWorkbooks;
using Microsoft.ConnectedWorkbooks.Internal;
using Microsoft.ConnectedWorkbooks.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ConnectedWorkbooks.Tests;

[TestClass]
public sealed class WorkbookManagerTests
{
    private readonly WorkbookManager _manager = new();
    private const string CustomSheetName = "DataSheet";
    private const string CustomSheetPath = "xl/worksheets/customSheet.xml";
    private const string CustomSheetRelsPath = "xl/worksheets/_rels/customSheet.xml.rels";
    private const string CustomTableName = "SalesTable";
    private const string CustomTablePath = "xl/tables/customTable.xml";
    private const string CustomTableRange = "C3:D5";

    [TestMethod]
    public async Task GeneratesWorkbookWithMashupAndTable()
    {
        var queryBody = @"let
    Source = Kusto.Contents(""https://help.kusto.windows.net"", ""Samples"", ""StormEvents"")
in
    Source";
        var query = new QueryInfo(queryBody, "DataAgentQuery", refreshOnOpen: true);
        var grid = new Grid(new List<IReadOnlyList<object?>>
        {
            new List<object?> { "City", "Count" },
            new List<object?> { "Seattle", 42 },
            new List<object?> { "London", 12 }
        }, new GridConfig { PromoteHeaders = true, AdjustColumnNames = true });

        var bytes = await _manager.GenerateSingleQueryWorkbookAsync(query, grid);
        Assert.IsTrue(bytes.Length > 0, "The generated workbook should not be empty.");

        using var archiveStream = new MemoryStream(bytes);
        using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read);

        var connections = ReadEntry(archive, "xl/connections.xml");
        StringAssert.Contains(connections, "DataAgentQuery");

        var sharedStrings = ReadEntry(archive, "xl/sharedStrings.xml");
        StringAssert.Contains(sharedStrings, "DataAgentQuery");

        var sheet = ReadEntry(archive, "xl/worksheets/sheet1.xml");
        StringAssert.Contains(sheet, "Seattle");
        StringAssert.Contains(sheet, "London");

        var table = ReadEntry(archive, "xl/tables/table1.xml");
        StringAssert.Contains(table, "City");
        StringAssert.Contains(table, "Count");

        var mashupXml = ReadEntry(archive, "customXml/item1.xml");
        StringAssert.Contains(mashupXml, "DataMashup");
        var root = XDocument.Parse(mashupXml).Root ?? throw new AssertFailedException("DataMashup XML root was missing.");
        var sectionContent = ExtractSection1m(root.Value.Trim());
        StringAssert.Contains(sectionContent, "DataAgentQuery");
    }

    [TestMethod]
    public async Task GeneratesTableWorkbookFromGrid()
    {
        var grid = new Grid(new List<IReadOnlyList<object?>>
        {
            new List<object?> { "Product", "Quantity", "Price" },
            new List<object?> { "Apples", 5, 1.25 },
            new List<object?> { "Bananas", 8, 0.99 }
        }, new GridConfig { PromoteHeaders = true });

        var bytes = await _manager.GenerateTableWorkbookFromGridAsync(grid);
        Assert.IsTrue(bytes.Length > 0, "The generated workbook should not be empty.");

        using var archiveStream = new MemoryStream(bytes);
        using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read);

        var tableXml = ReadEntry(archive, "xl/tables/table1.xml");
        StringAssert.Contains(tableXml, "Product");
        StringAssert.Contains(tableXml, "Quantity");
        StringAssert.Contains(tableXml, "Price");

        var sheetXml = ReadEntry(archive, "xl/worksheets/sheet1.xml");
        StringAssert.Contains(sheetXml, "Apples");
        StringAssert.Contains(sheetXml, "Bananas");
    }

    [TestMethod]
    public void RejectsInvalidQueryName()
    {
        var query = new QueryInfo("let Source = 1 in Source", "Invalid.Name", refreshOnOpen: false);
        Assert.ThrowsException<ArgumentException>(() => _manager.GenerateSingleQueryWorkbookAsync(query).GetAwaiter().GetResult());
    }

    [TestMethod]
    public async Task RequiresTemplateSettingsWhenDefaultsMissing()
    {
        var template = await CreateCustomTemplateAsync();
        var grid = new Grid(new List<IReadOnlyList<object?>>
        {
            new List<object?> { "Col1", "Col2" },
            new List<object?> { "A", "B" }
        }, new GridConfig { PromoteHeaders = true });

        var config = new FileConfiguration { TemplateBytes = template };

        await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => _manager.GenerateTableWorkbookFromGridAsync(grid, config));
    }

    [TestMethod]
    public async Task GeneratesTableWorkbookWithTemplateSettings()
    {
        var template = await CreateCustomTemplateAsync();
        var grid = new Grid(new List<IReadOnlyList<object?>>
        {
            new List<object?> { "Product", "Qty" },
            new List<object?> { "Apples", 5 },
            new List<object?> { "Bananas", 3 }
        }, new GridConfig { PromoteHeaders = true });

        var config = new FileConfiguration
        {
            TemplateBytes = template,
            TemplateSettings = new TemplateSettings
            {
                SheetName = CustomSheetName,
                TableName = CustomTableName
            }
        };

        var bytes = await _manager.GenerateTableWorkbookFromGridAsync(grid, config);

        using var archiveStream = new MemoryStream(bytes);
        using var archive = new ZipArchive(archiveStream, ZipArchiveMode.Read);

        var sheetXml = ReadEntry(archive, CustomSheetPath);
        var sheetDoc = XDocument.Parse(sheetXml);
        var sheetNs = sheetDoc.Root?.Name.Namespace ?? XNamespace.None;
        var firstCell = sheetDoc.Descendants(sheetNs + "c").First();
        Assert.AreEqual("C3", firstCell.Attribute("r")?.Value, "Table data was not written at the expected starting cell.");

        var tableXml = ReadEntry(archive, CustomTablePath);
        StringAssert.Contains(tableXml, $"name=\"{CustomTableName}\"");
        StringAssert.Contains(tableXml, CustomTableRange, "Table reference was not updated.");
    }

    private static string ReadEntry(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? throw new AssertFailedException($"Entry '{path}' was not found in the workbook.");
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }

    private static async Task<byte[]> CreateCustomTemplateAsync()
    {
        var template = await EmbeddedTemplateLoader.LoadBlankTableTemplateAsync();
        using var stream = new MemoryStream();
        stream.Write(template, 0, template.Length);

        using (var zip = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            RenameEntry(zip, "xl/worksheets/sheet1.xml", CustomSheetPath);
            RenameEntry(zip, "xl/worksheets/_rels/sheet1.xml.rels", CustomSheetRelsPath);
            RenameEntry(zip, "xl/tables/table1.xml", CustomTablePath);

            MutateXmlEntry(zip, "xl/workbook.xml", doc =>
            {
                var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
                var sheet = doc.Descendants(ns + "sheet").First();
                sheet.SetAttributeValue("name", CustomSheetName);
            });

            MutateXmlEntry(zip, "xl/_rels/workbook.xml.rels", doc =>
            {
                var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
                var relationship = doc.Descendants(ns + "Relationship")
                    .First(node => node.Attribute("Target")?.Value?.EndsWith("worksheets/sheet1.xml", StringComparison.OrdinalIgnoreCase) == true);
                relationship.SetAttributeValue("Target", "worksheets/customSheet.xml");
            });

            MutateXmlEntry(zip, CustomSheetRelsPath, doc =>
            {
                var ns = doc.Root?.Name.Namespace ?? XNamespace.None;
                foreach (var relationship in doc.Descendants(ns + "Relationship"))
                {
                    var target = relationship.Attribute("Target")?.Value;
                    if (target != null && target.EndsWith("../tables/table1.xml", StringComparison.OrdinalIgnoreCase))
                    {
                        relationship.SetAttributeValue("Target", "../tables/customTable.xml");
                    }
                }
            });

            MutateXmlEntry(zip, CustomTablePath, doc =>
            {
                doc.Root?.SetAttributeValue("name", CustomTableName);
                doc.Root?.SetAttributeValue("displayName", CustomTableName);
                doc.Root?.SetAttributeValue("ref", CustomTableRange);
            });
        }

        return stream.ToArray();
    }

    private static void RenameEntry(ZipArchive zip, string originalName, string newName)
    {
        var entry = zip.GetEntry(originalName) ?? throw new AssertFailedException($"Entry '{originalName}' not found in the workbook template.");
        using var buffer = new MemoryStream();
        using (var source = entry.Open())
        {
            source.CopyTo(buffer);
        }

        entry.Delete();
        var newEntry = zip.CreateEntry(newName);
        using var target = newEntry.Open();
        buffer.Position = 0;
        buffer.CopyTo(target);
    }

    private static void MutateXmlEntry(ZipArchive zip, string path, Action<XDocument> mutate)
    {
        var entry = zip.GetEntry(path) ?? throw new AssertFailedException($"Entry '{path}' not found in the workbook template.");
        string xml;
        using (var reader = new StreamReader(entry.Open(), Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
        {
            xml = reader.ReadToEnd();
        }

        var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
        mutate(doc);

        entry.Delete();
        var newEntry = zip.CreateEntry(path);
        using var writer = new StreamWriter(newEntry.Open(), new UTF8Encoding(false));
        writer.Write(doc.ToString(SaveOptions.DisableFormatting));
    }

    private static string ExtractSection1m(string dataMashupBase64)
    {
        var bytes = Convert.FromBase64String(dataMashupBase64);
        using var memory = new MemoryStream(bytes);
        using var binaryReader = new BinaryReader(memory);
        binaryReader.ReadBytes(4); // version header
        var packageSize = binaryReader.ReadInt32();
        var packageBytes = binaryReader.ReadBytes(packageSize);

        using var packageStream = new MemoryStream(packageBytes);
        using var packageZip = new ZipArchive(packageStream, ZipArchiveMode.Read);
        var entry = packageZip.GetEntry("Formulas/Section1.m") ?? throw new AssertFailedException("Section1.m was not found in the mashup package.");
        using var entryStream = entry.Open();
        using var reader = new StreamReader(entryStream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
    }
}

