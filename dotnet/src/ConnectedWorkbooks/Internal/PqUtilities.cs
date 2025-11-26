// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class PqUtilities
{
    public static (string Path, string Base64) GetDataMashup(ExcelArchive archive)
    {
        foreach (var entryPath in archive.EnumerateEntries(WorkbookConstants.CustomXmlFolder))
        {
            var match = WorkbookConstants.CustomXmlItemRegex.Match(entryPath);
            if (!match.Success)
            {
                continue;
            }

            var bytes = archive.ReadBytes(entryPath);
            var xml = XmlEncodingHelper.DecodeToString(bytes).TrimStart('\uFEFF');
            var doc = XDocument.Parse(xml, LoadOptions.PreserveWhitespace);
            if (!string.Equals(doc.Root?.Name.NamespaceName, WorkbookConstants.DataMashupNamespace, StringComparison.Ordinal))
            {
                continue;
            }

            var base64 = doc.Root?.Value ?? throw new InvalidOperationException("DataMashup element was empty.");
            return (entryPath, base64);
        }

        throw new InvalidOperationException("DataMashup XML was not found in the workbook template.");
    }

    public static void SetDataMashup(ExcelArchive archive, string path, string base64)
    {
        var xml = $"<?xml version=\"1.0\" encoding=\"utf-16\"?><DataMashup xmlns=\"{WorkbookConstants.DataMashupNamespace}\">{base64}</DataMashup>";
        var encoded = Encoding.Unicode.GetBytes("\uFEFF" + xml);
        archive.WriteBytes(path, encoded);
    }

    public static void ValidateQueryName(string queryName)
    {
        if (string.IsNullOrWhiteSpace(queryName))
        {
            throw new ArgumentException("Query name cannot be empty.", nameof(queryName));
        }

        if (queryName.Length > WorkbookConstants.MaxQueryLength)
        {
            throw new ArgumentException($"Query names are limited to {WorkbookConstants.MaxQueryLength} characters.", nameof(queryName));
        }

        if (queryName.Any(ch => ch == '"' || ch == '.' || char.IsControl(ch)))
        {
            throw new ArgumentException("Query names cannot contain periods, quotes, or control characters.", nameof(queryName));
        }
    }
}

