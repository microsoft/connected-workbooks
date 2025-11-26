// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Generates Power Query (M) documents used by the workbook templates.
/// </summary>
internal static class PowerQueryGenerator
{
    /// <summary>
    /// Creates a Section1.m document that exposes a single shared query with the supplied body.
    /// </summary>
    /// <param name="queryName">Name of the query to generate.</param>
    /// <param name="queryBody">M script that defines the query.</param>
    /// <returns>The complete M document ready to embed.</returns>
    public static string GenerateSingleQueryMashup(string queryName, string queryBody)
    {
        return $"section Section1;\n\nshared #\"{queryName}\" = \n{queryBody};";
    }
}

