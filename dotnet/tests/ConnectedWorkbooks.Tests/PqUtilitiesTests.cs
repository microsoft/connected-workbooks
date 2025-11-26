// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Text;
using Microsoft.ConnectedWorkbooks.Internal;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ConnectedWorkbooks.Tests;

[TestClass]
public sealed class PqUtilitiesTests
{
    [TestMethod]
    public async Task GetDataMashupHandlesUtf16LittleEndianBom()
    {
        await AssertDataMashupRoundtripAsync(Encoding.Unicode);
    }

    [TestMethod]
    public async Task GetDataMashupHandlesUtf16BigEndianBom()
    {
        await AssertDataMashupRoundtripAsync(Encoding.BigEndianUnicode);
    }

    [TestMethod]
    public async Task GetDataMashupHandlesUtf8Bom()
    {
        await AssertDataMashupRoundtripAsync(new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
    }

    private static async Task AssertDataMashupRoundtripAsync(Encoding encoding)
    {
        var template = await EmbeddedTemplateLoader.LoadSimpleQueryTemplateAsync();
        using var archive = ExcelArchive.Load(template);
        var base64 = Convert.ToBase64String(Guid.NewGuid().ToByteArray());
        WriteDataMashup(archive, base64, encoding);

        var (_, decodedBase64) = PqUtilities.GetDataMashup(archive);
        Assert.AreEqual(base64, decodedBase64, $"DataMashup decoding failed for encoding {encoding.WebName}.");
    }

    private static void WriteDataMashup(ExcelArchive archive, string base64, Encoding encoding)
    {
        var path = LocateDataMashupEntry(archive);
        var xml = $"<?xml version=\"1.0\" encoding=\"{encoding.WebName}\"?><DataMashup xmlns=\"{WorkbookConstants.DataMashupNamespace}\">{base64}</DataMashup>";
        var preamble = encoding.GetPreamble();
        var payload = encoding.GetBytes(xml);
        var buffer = new byte[preamble.Length + payload.Length];
        preamble.CopyTo(buffer, 0);
        payload.CopyTo(buffer, preamble.Length);
        archive.WriteBytes(path, buffer);
    }

    private static string LocateDataMashupEntry(ExcelArchive archive)
    {
        foreach (var entryPath in archive.EnumerateEntries(WorkbookConstants.CustomXmlFolder))
        {
            if (!WorkbookConstants.CustomXmlItemRegex.IsMatch(entryPath))
            {
                continue;
            }

            var xml = archive.ReadText(entryPath);
            var doc = System.Xml.Linq.XDocument.Parse(xml);
            if (string.Equals(doc.Root?.Name.LocalName, "DataMashup", StringComparison.Ordinal))
            {
                return entryPath;
            }
        }

        throw new AssertFailedException("DataMashup entry was not found in the template.");
    }
}
