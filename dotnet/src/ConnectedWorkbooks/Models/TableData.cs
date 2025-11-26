// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Normalized representation of tabular data that can be written into the workbook.
/// </summary>
/// <param name="ColumnNames">Column headers that appear in the table.</param>
/// <param name="Rows">Table rows aligned with <paramref name="ColumnNames"/>.</param>
public sealed record TableData(IReadOnlyList<string> ColumnNames, IReadOnlyList<IReadOnlyList<string>> Rows)
{
}

