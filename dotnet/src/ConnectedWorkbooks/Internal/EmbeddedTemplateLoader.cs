// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Reflection;

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Provides access to the workbook templates embedded in the assembly resources.
/// </summary>
internal static class EmbeddedTemplateLoader
{
    private const string ResourcePrefix = "ConnectedWorkbooks.Templates.";
    private const string SimpleQueryTemplateResource = ResourcePrefix + "SIMPLE_QUERY_WORKBOOK_TEMPLATE.xlsx";
    private const string BlankTableTemplateResource = ResourcePrefix + "SIMPLE_BLANK_TABLE_TEMPLATE.xlsx";
    private static readonly Lazy<byte[]> SimpleQueryTemplate = new(() => LoadTemplateFromManifest(SimpleQueryTemplateResource));
    private static readonly Lazy<byte[]> BlankTableTemplate = new(() => LoadTemplateFromManifest(BlankTableTemplateResource));

    /// <summary>
    /// Returns the cached bytes for the single-query workbook template.
    /// </summary>
    public static byte[] LoadSimpleQueryTemplate() =>
        SimpleQueryTemplate.Value;

    /// <summary>
    /// Returns the cached bytes for the blank table workbook template.
    /// </summary>
    public static byte[] LoadBlankTableTemplate() =>
        BlankTableTemplate.Value;

    /// <summary>
    /// Reads an embedded resource stream and returns its contents as a byte array.
    /// </summary>
    /// <param name="resourceName">Fully qualified manifest resource name.</param>
    /// <returns>The resource contents.</returns>
    private static byte[] LoadTemplateFromManifest(string resourceName)
    {
        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException($"Unable to locate embedded template '{resourceName}'.");
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }
}

