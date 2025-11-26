// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Controls how incoming grid data should be interpreted when converted into an Excel table.
/// </summary>
public sealed record GridConfig
{
    /// <summary>
    /// Treat the first row of <see cref="Grid.Data"/> as the header row.
    /// </summary>
    public bool PromoteHeaders { get; init; } = true;

    /// <summary>
    /// Automatically fix duplicate/blank headers by appending numeric suffixes.
    /// </summary>
    public bool AdjustColumnNames { get; init; } = true;
}

