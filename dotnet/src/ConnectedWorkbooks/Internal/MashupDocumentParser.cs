// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class MashupDocumentParser
{
    public static string ReplaceSingleQuery(string base64, string queryName, string queryMashupDocument)
    {
        var buffer = Convert.FromBase64String(base64);
        var reader = new ArrayReader(buffer);
        var versionBytes = reader.ReadBytes(4);
        var packageSize = reader.ReadInt32();
        var packageOpc = reader.ReadBytes(packageSize);
        var permissionsSize = reader.ReadInt32();
        var permissions = reader.ReadBytes(permissionsSize);
        var metadataSize = reader.ReadInt32();
        var metadataBytes = reader.ReadBytes(metadataSize);
        var endBuffer = reader.ReadToEnd();

        var newPackage = EditSingleQueryPackage(packageOpc, queryMashupDocument);
        var newMetadata = EditSingleQueryMetadata(metadataBytes, queryName);

        var finalBytes = ByteHelpers.Concat(
            versionBytes,
            ByteHelpers.GetInt32Bytes(newPackage.Length),
            newPackage,
            ByteHelpers.GetInt32Bytes(permissionsSize),
            permissions,
            ByteHelpers.GetInt32Bytes(newMetadata.Length),
            newMetadata,
            endBuffer);

        return Convert.ToBase64String(finalBytes);
    }

    private static byte[] EditSingleQueryPackage(byte[] packageOpc, string queryMashupDocument)
    {
        using var packageStream = new MemoryStream();
        packageStream.Write(packageOpc, 0, packageOpc.Length);
        packageStream.Position = 0;
        using var zip = new ZipArchive(packageStream, ZipArchiveMode.Update, leaveOpen: true);
        var entry = zip.GetEntry(WorkbookConstants.Section1mPath)
            ?? throw new InvalidOperationException("Formula section was not found in the Power Query package.");

        using (var entryStream = entry.Open())
        using (var writer = new StreamWriter(entryStream, new UTF8Encoding(false), leaveOpen: true))
        {
            entryStream.SetLength(0);
            writer.Write(queryMashupDocument);
            writer.Flush();
        }

        zip.Dispose();
        return packageStream.ToArray();
    }

    private static byte[] EditSingleQueryMetadata(byte[] metadataBytes, string queryName)
    {
        var reader = new ArrayReader(metadataBytes);
        var metadataVersion = reader.ReadBytes(4);
        var metadataXmlSize = reader.ReadInt32();
        var metadataXmlBytes = reader.ReadBytes(metadataXmlSize);
        var endBuffer = reader.ReadToEnd();

        var metadataXmlString = Encoding.UTF8.GetString(metadataXmlBytes).TrimStart('\uFEFF');
        XDocument metadataDoc;
        try
        {
            metadataDoc = XDocument.Parse(metadataXmlString, LoadOptions.PreserveWhitespace);
        }
        catch (Exception ex)
        {
            var preview = Convert.ToHexString(metadataXmlBytes.AsSpan(0, Math.Min(metadataXmlBytes.Length, 64)));
            throw new InvalidOperationException($"Failed to parse metadata XML. Hex preview: {preview}", ex);
        }
        UpdateItemPaths(metadataDoc, queryName);
        UpdateEntries(metadataDoc);

        var newMetadataXml = Encoding.UTF8.GetBytes(metadataDoc.ToString(SaveOptions.DisableFormatting));
        return ByteHelpers.Concat(
            metadataVersion,
            ByteHelpers.GetInt32Bytes(newMetadataXml.Length),
            newMetadataXml,
            endBuffer);
    }

    private static void UpdateItemPaths(XDocument doc, string queryName)
    {
        if (doc.Root is null)
        {
            return;
        }

        foreach (var itemPathElement in doc.Descendants().Where(e => e.Name.LocalName == XmlNames.Elements.ItemPath))
        {
            var content = itemPathElement.Value;
            if (!content.Contains("Section1/", StringComparison.Ordinal))
            {
                continue;
            }

            var parts = content.Split('/', StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 2)
            {
                continue;
            }

            parts[1] = Uri.EscapeDataString(queryName);
            itemPathElement.Value = string.Join('/', parts);
        }
    }

    private static void UpdateEntries(XDocument doc)
    {
        var now = DateTime.UtcNow.ToString("o", System.Globalization.CultureInfo.InvariantCulture);
        var lastUpdatedValue = ($"d{now}").Replace("Z", "0000Z", StringComparison.Ordinal);

        foreach (var entry in doc.Descendants().Where(e => e.Name.LocalName == XmlNames.Elements.Entry))
        {
            var typeValue = entry.Attribute(XmlNames.Attributes.Type)?.Value;
            if (string.Equals(typeValue, "ResultType", StringComparison.Ordinal))
            {
                entry.SetAttributeValue(XmlNames.Attributes.Value, "sTable");
            }
            else if (string.Equals(typeValue, "FillLastUpdated", StringComparison.Ordinal))
            {
                entry.SetAttributeValue(XmlNames.Attributes.Value, lastUpdatedValue);
            }
        }
    }
}

