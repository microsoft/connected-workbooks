// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System;
using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Handles low-level editing of the Power Query (DataMashup) payload inside workbook templates.
/// </summary>
internal static class MashupDocumentParser
{
    /// <summary>
    /// Replaces the contents of the Section1.m formula with the provided mashup and updates metadata accordingly.
    /// </summary>
    /// <param name="base64">Original DataMashup payload encoded as Base64.</param>
    /// <param name="queryName">Friendly query name that should be referenced throughout metadata.</param>
    /// <param name="queryMashupDocument">New M document to write into Section1.m.</param>
    /// <returns>A base64 string with the updated mashup payload.</returns>
    public static string ReplaceSingleQuery(string base64, string queryName, string queryMashupDocument)
    {
        var buffer = Convert.FromBase64String(base64);
        var reader = new ArrayReader(buffer);
        var versionBytes = reader.ReadMemory(4);
        var packageSize = reader.ReadInt32();
        var packageOpc = reader.ReadMemory(packageSize);
        var permissionsSize = reader.ReadInt32();
        var permissions = reader.ReadMemory(permissionsSize);
        var metadataSize = reader.ReadInt32();
        var metadataBytes = reader.ReadMemory(metadataSize);
        var endBuffer = reader.ReadToEnd();

        var newPackage = EditSingleQueryPackage(packageOpc.Span, queryMashupDocument);
        var newMetadata = EditSingleQueryMetadata(metadataBytes, queryName);

        var totalLength = versionBytes.Length
            + sizeof(int)
            + newPackage.Length
            + sizeof(int)
            + permissions.Length
            + sizeof(int)
            + newMetadata.Length
            + endBuffer.Length;

        var finalBytes = new byte[totalLength];
        var writer = new SpanWriter(finalBytes);
        writer.WriteBytes(versionBytes.Span);
        writer.WriteLength(newPackage.Length);
        writer.WriteBytes(newPackage);
        writer.WriteLength(permissionsSize);
        writer.WriteBytes(permissions.Span);
        writer.WriteLength(newMetadata.Length);
        writer.WriteBytes(newMetadata);
        writer.WriteBytes(endBuffer.Span);

        return Convert.ToBase64String(finalBytes);
    }

    /// <summary>
    /// Extracts the first query name referenced in the template's metadata.
    /// </summary>
    /// <param name="base64">Base64-encoded DataMashup payload.</param>
    /// <returns>The query name referenced by Section1 (defaults to Query1 when missing).</returns>
    public static string GetPrimaryQueryName(string base64)
    {
        var buffer = Convert.FromBase64String(base64);
        var reader = new ArrayReader(buffer);
        reader.ReadMemory(4); // version
        var packageSize = reader.ReadInt32();
        reader.ReadMemory(packageSize);
        var permissionsSize = reader.ReadInt32();
        reader.ReadMemory(permissionsSize);
        var metadataSize = reader.ReadInt32();
        var metadataBytes = reader.ReadMemory(metadataSize);

        var metadataReader = new ArrayReader(metadataBytes);
        metadataReader.ReadMemory(4); // metadata version
        var metadataXmlSize = metadataReader.ReadInt32();
        var metadataXmlBytes = metadataReader.ReadMemory(metadataXmlSize);

        var doc = ParseMetadataDocument(metadataXmlBytes);
        return ExtractQueryName(doc) ?? WorkbookConstants.DefaultQueryName;
    }

    private static byte[] EditSingleQueryPackage(ReadOnlySpan<byte> packageOpc, string queryMashupDocument)
    {
        using var packageStream = new MemoryStream();
        packageStream.Write(packageOpc);
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

    private static byte[] EditSingleQueryMetadata(ReadOnlyMemory<byte> metadataBytes, string queryName)
    {
        var reader = new ArrayReader(metadataBytes);
        var metadataVersion = reader.ReadMemory(4);
        var metadataXmlSize = reader.ReadInt32();
        var metadataXmlBytes = reader.ReadMemory(metadataXmlSize);
        var endBuffer = reader.ReadToEnd();

        var metadataDoc = ParseMetadataDocument(metadataXmlBytes);
        UpdateMetadataDocument(metadataDoc, queryName);

        var newMetadataXml = Encoding.UTF8.GetBytes(metadataDoc.ToString(SaveOptions.DisableFormatting));

        var totalLength = metadataVersion.Length + sizeof(int) + newMetadataXml.Length + endBuffer.Length;
        var buffer = new byte[totalLength];
        var writer = new SpanWriter(buffer);
        writer.WriteBytes(metadataVersion.Span);
        writer.WriteLength(newMetadataXml.Length);
        writer.WriteBytes(newMetadataXml);
        writer.WriteBytes(endBuffer.Span);

        return buffer;
    }

    private static XDocument ParseMetadataDocument(ReadOnlyMemory<byte> metadataXmlBytes)
    {
        var metadataXmlString = Encoding.UTF8.GetString(metadataXmlBytes.Span).TrimStart('\uFEFF');
        try
        {
            return XDocument.Parse(metadataXmlString, LoadOptions.PreserveWhitespace);
        }
        catch (Exception ex)
        {
            var preview = Convert.ToHexString(metadataXmlBytes.Span[..Math.Min(metadataXmlBytes.Length, 64)]);
            throw new InvalidOperationException($"Failed to parse metadata XML. Hex preview: {preview}", ex);
        }
    }

    private static string? ExtractQueryName(XDocument doc)
    {
        if (doc.Root is null)
        {
            return null;
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

            return Uri.UnescapeDataString(parts[1]);
        }

        return null;
    }

    private static void UpdateMetadataDocument(XDocument doc, string queryName)
    {
        if (!string.IsNullOrWhiteSpace(queryName))
        {
            RenameItemPaths(doc, queryName);
        }

        var now = DateTime.UtcNow.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss'.'fff", System.Globalization.CultureInfo.InvariantCulture);
        var lastUpdatedValue = $"d{now}0000Z";

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

    private static void RenameItemPaths(XDocument doc, string queryName)
    {
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

    private ref struct SpanWriter
    {
        private readonly Span<byte> _destination;
        private int _offset;

        public SpanWriter(Span<byte> destination)
        {
            _destination = destination;
            _offset = 0;
        }

        public void WriteBytes(ReadOnlySpan<byte> source)
        {
            if (source.Length == 0)
            {
                return;
            }

            source.CopyTo(_destination[_offset..]);
            _offset += source.Length;
        }

        public void WriteLength(int value)
        {
            BinaryPrimitives.WriteInt32LittleEndian(_destination.Slice(_offset, sizeof(int)), value);
            _offset += sizeof(int);
        }
    }
}

