// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Microsoft.ConnectedWorkbooks;
using Microsoft.ConnectedWorkbooks.Models;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ConnectedWorkbooks.Tests;

[TestClass]
public sealed class WorkbookManagerTests
{
    private readonly WorkbookManager _manager = new();

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
    public void RejectsInvalidQueryName()
    {
        var query = new QueryInfo("let Source = 1 in Source", "Invalid.Name", refreshOnOpen: false);
        Assert.ThrowsException<ArgumentException>(() => _manager.GenerateSingleQueryWorkbookAsync(query).GetAwaiter().GetResult());
    }

    private static string ReadEntry(ZipArchive archive, string path)
    {
        var entry = archive.GetEntry(path) ?? throw new AssertFailedException($"Entry '{path}' was not found in the workbook.");
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true);
        return reader.ReadToEnd();
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

