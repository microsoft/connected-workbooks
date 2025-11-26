// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Collections.Generic;
using System.Linq;
using Microsoft.ConnectedWorkbooks.Models;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class GridParser
{
    public static TableData Parse(Grid grid)
    {
        grid ??= new Grid(Array.Empty<IReadOnlyList<object?>>());
        var data = grid.Data?.Select(row => row.Select(value => value?.ToString() ?? string.Empty).ToArray()).ToList()
                   ?? new List<string[]> { Array.Empty<string>() };

        var promoteHeaders = grid.Config?.PromoteHeaders ?? false;
        var adjustColumns = grid.Config?.AdjustColumnNames ?? true;

        CorrectGrid(data, ref promoteHeaders);
        ValidateGrid(data, promoteHeaders, adjustColumns);

        string[] columnNames;
        if (promoteHeaders && adjustColumns)
        {
            columnNames = AdjustColumnNames(data[0]);
            data.RemoveAt(0);
        }
        else if (promoteHeaders)
        {
            columnNames = data[0];
            data.RemoveAt(0);
        }
        else
        {
            columnNames = Enumerable.Range(1, data[0].Length).Select(i => $"Column {i}").ToArray();
        }

        return new TableData(columnNames, data);
    }

    private static void CorrectGrid(IList<string[]> data, ref bool promoteHeaders)
    {
        if (data.Count == 0)
        {
            promoteHeaders = false;
            data.Add(new[] { string.Empty });
            return;
        }

        if (data[0].Length == 0)
        {
            data[0] = new[] { string.Empty };
        }

        var width = data[0].Length;
        for (var i = 0; i < data.Count; i++)
        {
            if (data[i].Length == 0)
            {
                data[i] = Enumerable.Repeat(string.Empty, width).ToArray();
            }
        }

        if (promoteHeaders && data.Count == 1)
        {
            data.Add(Enumerable.Repeat(string.Empty, width).ToArray());
        }
    }

    private static void ValidateGrid(IReadOnlyList<string[]> data, bool promoteHeaders, bool adjustColumns)
    {
        if (data.Count == 0 || data[0].Length == 0)
        {
            throw new InvalidOperationException("The provided grid is empty.");
        }

        if (data.Any(row => row.Length != data[0].Length))
        {
            throw new InvalidOperationException("The provided grid is not a rectangular MxN matrix.");
        }

        if (promoteHeaders && !adjustColumns)
        {
            if (data[0].Any(string.IsNullOrWhiteSpace))
            {
                throw new InvalidOperationException("Headers cannot be promoted when empty values exist.");
            }

            var uniqueCount = data[0].Select(name => name.ToLowerInvariant()).Distinct().Count();
            if (uniqueCount != data[0].Length)
            {
                throw new InvalidOperationException("Headers must be unique when column adjustments are disabled.");
            }
        }
    }

    private static string[] AdjustColumnNames(string[] columnNames)
    {
        var unique = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new string[columnNames.Length];
        for (var i = 0; i < columnNames.Length; i++)
        {
            var baseName = string.IsNullOrWhiteSpace(columnNames[i]) ? $"Column {i + 1}" : columnNames[i];
            var candidate = baseName;
            var suffix = 1;
            while (!unique.Add(candidate))
            {
                candidate = $"{baseName} ({suffix++})";
            }

            result[i] = candidate;
        }

        return result;
    }
}

