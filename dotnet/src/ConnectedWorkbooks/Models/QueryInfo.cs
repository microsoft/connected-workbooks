// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Describes a single Power Query definition that should be injected into the generated workbook.
/// </summary>
public sealed record QueryInfo
{
    /// <summary>
    /// Creates a new query description.
    /// </summary>
    /// <param name="queryMashup">The Power Query (M) text that defines the query.</param>
    /// <param name="queryName">Optional friendly name; defaults to <c>Query1</c> if omitted.</param>
    /// <param name="refreshOnOpen">Whether Excel should refresh the query automatically when the workbook opens.</param>
    public QueryInfo(string queryMashup, string? queryName = null, bool refreshOnOpen = true)
    {
        QueryMashup = string.IsNullOrWhiteSpace(queryMashup)
            ? throw new ArgumentException("Query mashup cannot be null or empty.", nameof(queryMashup))
            : queryMashup;

        QueryName = string.IsNullOrWhiteSpace(queryName) ? null : queryName;
        RefreshOnOpen = refreshOnOpen;
    }

    /// <summary>
    /// Gets the Power Query (M) script.
    /// </summary>
    public string QueryMashup { get; }

    /// <summary>
    /// Gets the friendly query name (or <c>null</c> to fall back to the default).
    /// </summary>
    public string? QueryName { get; }

    /// <summary>
    /// Gets a value indicating whether the query should refresh when the workbook opens.
    /// </summary>
    public bool RefreshOnOpen { get; }
}

