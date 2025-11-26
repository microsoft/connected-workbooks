// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Optional overrides used when supplying a custom workbook template.
/// </summary>
public sealed record TemplateSettings
{
    /// <summary>
    /// Optional table name override inside the custom template.
    /// </summary>
    public string? TableName { get; init; }

    /// <summary>
    /// Optional worksheet name override inside the custom template.
    /// </summary>
    public string? SheetName { get; init; }
}

