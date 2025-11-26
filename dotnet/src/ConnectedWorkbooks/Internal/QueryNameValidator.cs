// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System;

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Provides validation and normalization helpers for user-supplied query names.
/// </summary>
internal static class QueryNameValidator
{
    /// <summary>
    /// Normalizes a user-supplied query name, applying defaults and validation.
    /// </summary>
    /// <param name="candidate">The original query name provided by the caller.</param>
    /// <returns>A validated query name.</returns>
    /// <exception cref="ArgumentException">Thrown when the name does not meet requirements.</exception>
    public static string Resolve(string? candidate)
    {
        var effectiveName = string.IsNullOrWhiteSpace(candidate)
            ? WorkbookConstants.DefaultQueryName
            : candidate.Trim();

        Validate(effectiveName);
        return effectiveName;
    }

    /// <summary>
    /// Validates a query name using the same constraints as the TypeScript implementation.
    /// </summary>
    /// <param name="queryName">Name to validate.</param>
    /// <exception cref="ArgumentException">Thrown when the name does not meet requirements.</exception>
    public static void Validate(string queryName)
    {
        if (string.IsNullOrWhiteSpace(queryName))
        {
            throw new ArgumentException("Query name cannot be empty.", nameof(queryName));
        }

        if (queryName.Length > WorkbookConstants.MaxQueryLength)
        {
            throw new ArgumentException($"Query names are limited to {WorkbookConstants.MaxQueryLength} characters.", nameof(queryName));
        }

        foreach (var ch in queryName)
        {
            if (ch == '"' || ch == '.' || IsControlCharacter(ch))
            {
                throw new ArgumentException("Query name contains invalid characters.", nameof(queryName));
            }
        }
    }

    private static bool IsControlCharacter(char value)
    {
        return (value >= '\u0000' && value <= '\u001F')
            || (value >= '\u007F' && value <= '\u009F');
    }
}
