// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Optional document metadata used to stamp core properties in the generated workbook.
/// </summary>
public sealed record DocumentProperties
{
    public string? Title { get; init; }
    public string? Subject { get; init; }
    public string? Keywords { get; init; }
    public string? CreatedBy { get; init; }
    public string? Description { get; init; }
    public string? LastModifiedBy { get; init; }
    public string? Category { get; init; }
    public string? Revision { get; init; }
}

