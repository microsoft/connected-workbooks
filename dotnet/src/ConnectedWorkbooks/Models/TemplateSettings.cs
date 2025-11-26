// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Optional overrides used when supplying a custom workbook template.
/// </summary>
public sealed record TemplateSettings
{
    public string? TableName { get; init; }
    public string? SheetName { get; init; }
}

