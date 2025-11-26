// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using Microsoft.ConnectedWorkbooks.Internal;
using Microsoft.ConnectedWorkbooks.Models;

namespace Microsoft.ConnectedWorkbooks;

/// <summary>
/// Entry point for generating Connected Workbooks from .NET.
/// </summary>
public sealed class WorkbookManager
{
    /// <summary>
    /// Generates a workbook that contains a single Power Query backed by optional initial grid data.
    /// </summary>
    /// <param name="query">Information about the query to embed.</param>
    /// <param name="initialDataGrid">Optional grid whose data should seed the query table.</param>
    /// <param name="fileConfiguration">Optional template and document configuration overrides.</param>
    /// <returns>The generated workbook bytes.</returns>
    public byte[] GenerateSingleQueryWorkbook(
        QueryInfo query,
        Grid? initialDataGrid = null,
        FileConfiguration? fileConfiguration = null)
    {
        ArgumentNullException.ThrowIfNull(query);

        var templateBytes = fileConfiguration?.TemplateBytes
            ?? EmbeddedTemplateLoader.LoadSimpleQueryTemplate();
        var tableData = initialDataGrid is null ? null : GridParser.Parse(initialDataGrid);
        var effectiveQueryName = QueryNameValidator.Resolve(query.QueryName);

        using var archive = ExcelArchive.Load(templateBytes);
        var templateMetadata = TemplateMetadataResolver.Resolve(archive, fileConfiguration?.TemplateSettings);
        var editor = new WorkbookEditor(archive, fileConfiguration?.DocumentProperties, templateMetadata);
        editor.UpdatePowerQueryDocument(query.QueryMashup, effectiveQueryName);

        var connectionId = editor.UpdateConnections(effectiveQueryName, query.RefreshOnOpen);
        var sharedStringIndex = editor.UpdateSharedStrings(effectiveQueryName);
        editor.UpdateWorksheet(sharedStringIndex);
        editor.UpdateQueryTable(connectionId, query.RefreshOnOpen);
        if (tableData is not null)
        {
            editor.UpdateTableData(tableData);
        }
        editor.UpdateDocumentProperties();

        return archive.ToArray();
    }

    /// <summary>
    /// Generates a workbook that contains only a table populated from the supplied grid.
    /// </summary>
    /// <param name="grid">The source grid data.</param>
    /// <param name="fileConfiguration">Optional template/document overrides.</param>
    /// <returns>The generated workbook bytes.</returns>
    public byte[] GenerateTableWorkbookFromGrid(
        Grid grid,
        FileConfiguration? fileConfiguration = null)
    {
        ArgumentNullException.ThrowIfNull(grid);

        var templateBytes = fileConfiguration?.TemplateBytes
            ?? EmbeddedTemplateLoader.LoadBlankTableTemplate();
        var tableData = GridParser.Parse(grid);

        using var archive = ExcelArchive.Load(templateBytes);
        var templateMetadata = TemplateMetadataResolver.Resolve(archive, fileConfiguration?.TemplateSettings);
        var editor = new WorkbookEditor(archive, fileConfiguration?.DocumentProperties, templateMetadata);
        editor.UpdateTableData(tableData);
        editor.UpdateDocumentProperties();

        return archive.ToArray();
    }
}
