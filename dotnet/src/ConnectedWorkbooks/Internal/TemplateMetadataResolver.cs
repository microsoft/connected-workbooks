// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Linq;
using System.Xml.Linq;
using Microsoft.ConnectedWorkbooks.Models;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal sealed record TemplateMetadata(string WorksheetPath, string TablePath, (int Row, int Column) TableStart);

internal static class TemplateMetadataResolver
{
    private static readonly XNamespace RelationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace OfficeRelationshipsNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static TemplateMetadata Resolve(ExcelArchive archive, TemplateSettings? templateSettings)
    {
        var worksheetPath = ResolveWorksheetPath(archive, templateSettings?.SheetName);
        var (tablePath, tableStart) = ResolveTablePath(archive, templateSettings?.TableName);
        return new TemplateMetadata(worksheetPath, tablePath, tableStart);
    }

    private static string ResolveWorksheetPath(ExcelArchive archive, string? sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            if (!archive.EntryExists(WorkbookConstants.DefaultSheetPath))
            {
                throw new InvalidOperationException("The workbook template does not contain the default worksheet 'xl/worksheets/sheet1.xml'. Provide FileConfiguration.TemplateSettings.SheetName to indicate the target sheet.");
            }

            return WorkbookConstants.DefaultSheetPath;
        }

        var workbookXml = archive.ReadText(WorkbookConstants.WorkbookXmlPath);
        var workbookDoc = XDocument.Parse(workbookXml, LoadOptions.PreserveWhitespace);
        var workbookNs = workbookDoc.Root?.Name.Namespace ?? XNamespace.None;
        var sheetElement = workbookDoc
            .Descendants(workbookNs + "sheet")
            .FirstOrDefault(node => string.Equals(node.Attribute("name")?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"Worksheet '{sheetName}' was not found in the workbook template.");

        var relationshipId = sheetElement.Attribute(OfficeRelationshipsNamespace + "id")?.Value
            ?? throw new InvalidOperationException($"Worksheet '{sheetName}' is missing the relationship id attribute.");

        var relsXml = archive.ReadText(WorkbookConstants.WorkbookRelsPath);
        var relsDoc = XDocument.Parse(relsXml, LoadOptions.PreserveWhitespace);
        var relationship = relsDoc
            .Descendants(RelationshipsNamespace + XmlNames.Elements.Relationship)
            .FirstOrDefault(node => string.Equals(node.Attribute("Id")?.Value, relationshipId, StringComparison.Ordinal))
            ?? throw new InvalidOperationException($"Relationship '{relationshipId}' referenced by worksheet '{sheetName}' was not found.");

        var target = relationship.Attribute("Target")?.Value
            ?? throw new InvalidOperationException($"Relationship '{relationshipId}' does not contain a target path.");

        return NormalizePath(target);
    }

    private static (string Path, (int Row, int Column) Start) ResolveTablePath(ExcelArchive archive, string? tableName)
    {
        if (!string.IsNullOrWhiteSpace(tableName))
        {
            foreach (var entryPath in archive.EnumerateEntries(WorkbookConstants.TablesFolder))
            {
                var metadata = ReadTableMetadata(archive.ReadText(entryPath));
                if (string.Equals(metadata.Name, tableName, StringComparison.OrdinalIgnoreCase))
                {
                    return (entryPath, metadata.Start);
                }
            }

            throw new InvalidOperationException($"Table '{tableName}' was not found in the workbook template.");
        }

        if (!archive.EntryExists(WorkbookConstants.DefaultTablePath))
        {
            throw new InvalidOperationException("The workbook template does not contain the default table 'xl/tables/table1.xml'. Provide FileConfiguration.TemplateSettings.TableName to indicate which table to target.");
        }

        var defaultMetadata = ReadTableMetadata(archive.ReadText(WorkbookConstants.DefaultTablePath));
        return (WorkbookConstants.DefaultTablePath, defaultMetadata.Start);
    }

    private static (string Name, (int Row, int Column) Start) ReadTableMetadata(string tableXml)
    {
        var doc = XDocument.Parse(tableXml, LoadOptions.PreserveWhitespace);
        var name = doc.Root?.Attribute("name")?.Value
            ?? throw new InvalidOperationException("Table definition is missing the 'name' attribute.");
        var reference = doc.Root?.Attribute("ref")?.Value ?? "A1";
        var cleanedReference = reference.Replace("$", string.Empty, StringComparison.Ordinal);
        var start = CellReferenceHelper.GetStartPosition(cleanedReference);
        return (name, start);
    }

    private static string NormalizePath(string target)
    {
        var normalized = target.Replace('\\', '/');
        if (normalized.StartsWith("/", StringComparison.Ordinal))
        {
            return $"xl{normalized}";
        }

        return $"xl/{normalized}";
    }
}
