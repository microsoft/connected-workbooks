// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Describes a single Power Query definition that should be injected into the generated workbook.
/// </summary>
public sealed record QueryInfo
{
    public QueryInfo(string queryMashup, string? queryName = null, bool refreshOnOpen = true)
    {
        QueryMashup = string.IsNullOrWhiteSpace(queryMashup)
            ? throw new ArgumentException("Query mashup cannot be null or empty.", nameof(queryMashup))
            : queryMashup;

        QueryName = string.IsNullOrWhiteSpace(queryName) ? null : queryName;
        RefreshOnOpen = refreshOnOpen;
    }

    public string QueryMashup { get; }

    public string? QueryName { get; }

    public bool RefreshOnOpen { get; }
}

