// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Utility methods for converting between Excel-style cell references and numeric coordinates.
/// </summary>
internal static class CellReferenceHelper
{
    /// <summary>
    /// Returns the zero-based row/column tuple that represents the starting cell in the reference.
    /// </summary>
    /// <param name="reference">A cell reference such as "A1" or "A1:B5".</param>
    /// <returns>A tuple containing one-based row and column coordinates.</returns>
    public static (int Row, int Column) GetStartPosition(string reference)
    {
        // Reference format "A1" or "A1:B5"; we only care about the first cell
        var start = reference.Split(':')[0];
        var letters = new string(start.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(start.SkipWhile(char.IsLetter).ToArray());
        var column = ColumnNameToNumber(letters);
        var row = int.TryParse(digits, out var parsedRow) ? parsedRow : 1;
        return (row, column);
    }

    /// <summary>
    /// Converts a zero-based column index into its Excel column name.
    /// </summary>
    /// <param name="columnIndex">Zero-based column index.</param>
    /// <returns>The Excel column label (e.g. 0 -> "A").</returns>
    public static string ColumnNumberToName(int columnIndex)
    {
        columnIndex++; // zero-based to one-based
        var columnName = string.Empty;
        while (columnIndex > 0)
        {
            var remainder = (columnIndex - 1) % 26;
            columnName = (char)('A' + remainder) + columnName;
            columnIndex = (columnIndex - remainder - 1) / 26;
        }

        return columnName;
    }

    /// <summary>
    /// Builds a rectangular range reference given a starting coordinate and bounds.
    /// </summary>
    /// <param name="start">One-based starting position.</param>
    /// <param name="columnCount">Number of columns in the range.</param>
    /// <param name="rowCount">Number of rows in the range.</param>
    /// <returns>The Excel range reference spanning the requested area.</returns>
    public static string BuildReference((int Row, int Column) start, int columnCount, int rowCount)
    {
        var endColumnIndex = start.Column - 1 + columnCount;
        var endRow = start.Row - 1 + rowCount;
        var startRef = $"{ColumnNumberToName(start.Column - 1)}{start.Row}";
        var endRef = $"{ColumnNumberToName(endColumnIndex - 1)}{endRow}";
        return $"{startRef}:{endRef}";
    }

    /// <summary>
    /// Converts the provided range into an absolute reference with an optional sheet prefix.
    /// </summary>
    /// <param name="reference">The relative range reference.</param>
    /// <param name="sheetName">Optional sheet prefix (for example <c>'Sheet1'</c>).</param>
    /// <returns>The absolute equivalent (e.g. <c>'Sheet1'!$A$1:$B$2</c>).</returns>
    public static string WithAbsolute(string reference, string? sheetName = null)
    {
        var (row, column) = GetStartPosition(reference);
        var (endRow, endColumn) = GetEndPosition(reference);
        var prefix = string.IsNullOrEmpty(sheetName) ? string.Empty : $"{sheetName}!";
        return $"{prefix}${ColumnNumberToName(column - 1)}${row}:${ColumnNumberToName(endColumn - 1)}${endRow}";
    }

    private static (int Row, int Column) GetEndPosition(string reference)
    {
        var parts = reference.Split(':');
        var target = parts.Length == 2 ? parts[1] : parts[0];
        var letters = new string(target.TakeWhile(char.IsLetter).ToArray());
        var digits = new string(target.SkipWhile(char.IsLetter).ToArray());
        var column = ColumnNameToNumber(letters);
        var row = int.TryParse(digits, out var parsedRow) ? parsedRow : 1;
        return (row, column);
    }

    private static int ColumnNameToNumber(string columnName)
    {
        if (string.IsNullOrWhiteSpace(columnName))
        {
            return 1;
        }

        var result = 0;
        foreach (var ch in columnName.ToUpperInvariant())
        {
            result = result * 26 + (ch - 'A' + 1);
        }

        return result;
    }
}

