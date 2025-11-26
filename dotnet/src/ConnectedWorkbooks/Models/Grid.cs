// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Models;

/// <summary>
/// Simple 2D grid abstraction used for data ingestion.
/// </summary>
public sealed record Grid(IReadOnlyList<IReadOnlyList<object?>> Data, GridConfig? Config = null);

