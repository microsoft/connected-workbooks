// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Simple 2D grid abstraction used for data ingestion.
/// </summary>
/// <param name="Data">Rows and columns that represent the dataset.</param>
/// <param name="Config">Optional configuration that influences how the grid is interpreted.</param>
public sealed record Grid(IReadOnlyList<IReadOnlyList<object?>> Data, GridConfig? Config = null);

