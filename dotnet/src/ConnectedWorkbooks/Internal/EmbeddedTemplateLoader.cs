// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Reflection;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class EmbeddedTemplateLoader
{
    private const string ResourcePrefix = "ConnectedWorkbooks.Templates.";
    private const string SimpleQueryTemplateResource = ResourcePrefix + "SIMPLE_QUERY_WORKBOOK_TEMPLATE.xlsx";
    private const string BlankTableTemplateResource = ResourcePrefix + "SIMPLE_BLANK_TABLE_TEMPLATE.xlsx";

    public static byte[] LoadSimpleQueryTemplate() => LoadTemplate(SimpleQueryTemplateResource);

    public static byte[] LoadBlankTableTemplate() => LoadTemplate(BlankTableTemplateResource);

    private static byte[] LoadTemplate(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Unable to locate embedded template '{resourceName}'.");
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
}

