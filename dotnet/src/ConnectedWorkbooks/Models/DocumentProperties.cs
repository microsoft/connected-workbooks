// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Optional document metadata used to stamp core properties in the generated workbook.
/// </summary>
public sealed record DocumentProperties
{
    /// <summary>
    /// Value written to <c>dc:title</c>.
    /// </summary>
    public string? Title { get; init; }

    /// <summary>
    /// Value written to <c>dc:subject</c>.
    /// </summary>
    public string? Subject { get; init; }

    /// <summary>
    /// Optional keywords associated with the document.
    /// </summary>
    public string? Keywords { get; init; }

    /// <summary>
    /// Author recorded in <c>dc:creator</c>.
    /// </summary>
    public string? CreatedBy { get; init; }

    /// <summary>
    /// Long-form description of the workbook.
    /// </summary>
    public string? Description { get; init; }

    /// <summary>
    /// Value written to <c>cp:lastModifiedBy</c>.
    /// </summary>
    public string? LastModifiedBy { get; init; }

    /// <summary>
    /// Optional category classification.
    /// </summary>
    public string? Category { get; init; }

    /// <summary>
    /// Revision string assigned to <c>cp:revision</c>.
    /// </summary>
    public string? Revision { get; init; }
}

