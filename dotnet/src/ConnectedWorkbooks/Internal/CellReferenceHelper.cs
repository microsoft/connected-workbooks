// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Linq;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class CellReferenceHelper
{
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

    public static string BuildReference((int Row, int Column) start, int columnCount, int rowCount)
    {
        var endColumnIndex = start.Column - 1 + columnCount;
        var endRow = start.Row - 1 + rowCount;
        var startRef = $"{ColumnNumberToName(start.Column - 1)}{start.Row}";
        var endRef = $"{ColumnNumberToName(endColumnIndex - 1)}{endRow}";
        return $"{startRef}:{endRef}";
    }

    public static string WithAbsolute(string reference)
    {
        var (row, column) = GetStartPosition(reference);
        var (endRow, endColumn) = GetEndPosition(reference);
        return $"!${ColumnNumberToName(column - 1)}${row}:${ColumnNumberToName(endColumn - 1)}${endRow}";
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

