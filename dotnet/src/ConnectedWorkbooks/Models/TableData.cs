// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Normalized representation of tabular data that can be written into the workbook.
/// </summary>
public sealed record TableData(IReadOnlyList<string> ColumnNames, IReadOnlyList<IReadOnlyList<string>> Rows)
{
}

