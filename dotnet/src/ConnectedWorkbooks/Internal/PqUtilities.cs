// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Text;
using System.Xml.Linq;

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Helpers for interacting with the Power Query (DataMashup) parts inside workbook templates.
/// </summary>
internal static class PqUtilities
{
    /// <summary>
    /// Locates the DataMashup XML inside the workbook and returns its location and payload.
    /// </summary>
    /// <param name="archive">Workbook archive to inspect.</param>
    /// <returns>The entry path and base64 payload.</returns>
    public static (string Path, string Base64) GetDataMashup(ExcelArchive archive)
    {
        foreach (var entryPath in archive.EnumerateEntries(WorkbookConstants.CustomXmlFolder))
        {
            var match = WorkbookConstants.CustomXmlItemRegex.Match(entryPath);
            if (!match.Success)
            {
                continue;
            }

            var xml = archive.ReadText(entryPath);
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

    /// <summary>
    /// Writes the provided DataMashup payload back into the workbook.
    /// </summary>
    /// <param name="archive">Workbook archive to mutate.</param>
    /// <param name="path">Entry path returned by <see cref="GetDataMashup"/>.</param>
    /// <param name="base64">Base64-encoded payload that should replace the current content.</param>
    public static void SetDataMashup(ExcelArchive archive, string path, string base64)
    {
        var xml = $"<?xml version=\"1.0\" encoding=\"utf-16\"?><DataMashup xmlns=\"{WorkbookConstants.DataMashupNamespace}\">{base64}</DataMashup>";
        var encoded = Encoding.Unicode.GetBytes("\uFEFF" + xml);
        archive.WriteBytes(path, encoded);
    }

    /// <summary>
    /// Validates that a query name conforms to Excel's constraints.
    /// </summary>
    /// <param name="queryName">The user supplied query name.</param>
    /// <exception cref="ArgumentException">Thrown when the name violates naming rules.</exception>
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

