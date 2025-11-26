// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Reflection;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class EmbeddedTemplateLoader
{
    private const string ResourcePrefix = "ConnectedWorkbooks.Templates.";
    private const string SimpleQueryTemplateResource = ResourcePrefix + "SIMPLE_QUERY_WORKBOOK_TEMPLATE.xlsx";
    private const string BlankTableTemplateResource = ResourcePrefix + "SIMPLE_BLANK_TABLE_TEMPLATE.xlsx";

    public static Task<byte[]> LoadSimpleQueryTemplateAsync(CancellationToken cancellationToken = default) =>
        LoadTemplateAsync(SimpleQueryTemplateResource, cancellationToken);

    public static Task<byte[]> LoadBlankTableTemplateAsync(CancellationToken cancellationToken = default) =>
        LoadTemplateAsync(BlankTableTemplateResource, cancellationToken);

    private static async Task<byte[]> LoadTemplateAsync(string resourceName, CancellationToken cancellationToken)
    {
        var assembly = Assembly.GetExecutingAssembly();
        await using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Unable to locate embedded template '{resourceName}'.");
        await using var memory = new MemoryStream();
        await stream.CopyToAsync(memory, cancellationToken);
        return memory.ToArray();
    }
}

