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
    public async Task<byte[]> GenerateSingleQueryWorkbookAsync(
        QueryInfo query,
        Grid? initialDataGrid = null,
        FileConfiguration? fileConfiguration = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(query);

        var queryName = string.IsNullOrWhiteSpace(query.QueryName)
            ? "Query1"
            : query.QueryName!;

        PqUtilities.ValidateQueryName(queryName);
        var templateBytes = fileConfiguration?.TemplateBytes
            ?? await EmbeddedTemplateLoader.LoadSimpleQueryTemplateAsync(cancellationToken).ConfigureAwait(false);
        var tableData = initialDataGrid is null ? null : GridParser.Parse(initialDataGrid);

        using var archive = ExcelArchive.Load(templateBytes);
        var templateMetadata = TemplateMetadataResolver.Resolve(archive, fileConfiguration?.TemplateSettings);
        var editor = new WorkbookEditor(archive, fileConfiguration?.DocumentProperties, templateMetadata);
        var mashup = PowerQueryGenerator.GenerateSingleQueryMashup(queryName, query.QueryMashup);
        editor.UpdatePowerQueryDocument(queryName, mashup);
        var connectionId = editor.UpdateConnections(queryName, query.RefreshOnOpen);
        var sharedStringIndex = editor.UpdateSharedStrings(queryName);
        editor.UpdateWorksheet(sharedStringIndex);
        editor.UpdateQueryTable(connectionId, query.RefreshOnOpen);
        if (tableData is not null)
        {
            editor.UpdateTableData(tableData);
        }
        editor.UpdateDocumentProperties();

        return await Task.FromResult(archive.ToArray());
    }

    public async Task<byte[]> GenerateTableWorkbookFromGridAsync(
        Grid grid,
        FileConfiguration? fileConfiguration = null,
        CancellationToken cancellationToken = default)
    {
        ArgumentNullException.ThrowIfNull(grid);

        var templateBytes = fileConfiguration?.TemplateBytes
            ?? await EmbeddedTemplateLoader.LoadBlankTableTemplateAsync(cancellationToken).ConfigureAwait(false);
        var tableData = GridParser.Parse(grid);

        using var archive = ExcelArchive.Load(templateBytes);
        var templateMetadata = TemplateMetadataResolver.Resolve(archive, fileConfiguration?.TemplateSettings);
        var editor = new WorkbookEditor(archive, fileConfiguration?.DocumentProperties, templateMetadata);
        editor.UpdateTableData(tableData);
        editor.UpdateDocumentProperties();

        return await Task.FromResult(archive.ToArray());
    }
}
