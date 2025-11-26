// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class MashupDocumentParser
{
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
        var destination = finalBytes.AsSpan();
        var offset = 0;

        offset += Copy(versionBytes.Span, destination[offset..]);
        BinaryPrimitives.WriteInt32LittleEndian(destination.Slice(offset, sizeof(int)), newPackage.Length);
        offset += sizeof(int);
        offset += Copy(newPackage, destination[offset..]);
        BinaryPrimitives.WriteInt32LittleEndian(destination.Slice(offset, sizeof(int)), permissionsSize);
        offset += sizeof(int);
        offset += Copy(permissions.Span, destination[offset..]);
        BinaryPrimitives.WriteInt32LittleEndian(destination.Slice(offset, sizeof(int)), newMetadata.Length);
        offset += sizeof(int);
        offset += Copy(newMetadata, destination[offset..]);
        _ = Copy(endBuffer.Span, destination[offset..]);

        return Convert.ToBase64String(finalBytes);
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

        var metadataXmlString = Encoding.UTF8.GetString(metadataXmlBytes.Span).TrimStart('\uFEFF');
        XDocument metadataDoc;
        try
        {
            metadataDoc = XDocument.Parse(metadataXmlString, LoadOptions.PreserveWhitespace);
        }
        catch (Exception ex)
        {
            var preview = Convert.ToHexString(metadataXmlBytes.Span[..Math.Min(metadataXmlBytes.Length, 64)]);
            throw new InvalidOperationException($"Failed to parse metadata XML. Hex preview: {preview}", ex);
        }
        UpdateItemPaths(metadataDoc, queryName);
        UpdateEntries(metadataDoc);

        var newMetadataXml = Encoding.UTF8.GetBytes(metadataDoc.ToString(SaveOptions.DisableFormatting));

        var totalLength = metadataVersion.Length + sizeof(int) + newMetadataXml.Length + endBuffer.Length;
        var buffer = new byte[totalLength];
        var destination = buffer.AsSpan();
        var offset = 0;

        offset += Copy(metadataVersion.Span, destination[offset..]);
        BinaryPrimitives.WriteInt32LittleEndian(destination.Slice(offset, sizeof(int)), newMetadataXml.Length);
        offset += sizeof(int);
        offset += Copy(newMetadataXml, destination[offset..]);
        _ = Copy(endBuffer.Span, destination[offset..]);

        return buffer;
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
        var lastUpdatedValue = $"d{now}".Replace("Z", "0000Z", StringComparison.Ordinal);

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

    private static int Copy(ReadOnlySpan<byte> source, Span<byte> destination)
    {
        source.CopyTo(destination);
        return source.Length;
    }
}

