// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class PowerQueryGenerator
{
    public static string GenerateSingleQueryMashup(string queryName, string queryBody)
    {
        return $"section Section1;\n\nshared \"{queryName}\" = \n{queryBody};";
    }
}

