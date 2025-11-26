// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Optional knobs used when generating a workbook from the .NET implementation.
/// </summary>
public sealed record FileConfiguration
{
    /// <summary>
    /// When provided, the workbook will be generated using the supplied template bytes instead of the built-in one.
    /// </summary>
    public byte[]? TemplateBytes { get; init; }

    /// <summary>
    /// Document metadata that should be applied to <c>docProps/core.xml</c>.
    /// </summary>
    public DocumentProperties? DocumentProperties { get; init; }

    /// <summary>
    /// Fine grained instructions that help the generator locate the right sheet/table inside a custom template.
    /// </summary>
    public TemplateSettings? TemplateSettings { get; init; }
}

