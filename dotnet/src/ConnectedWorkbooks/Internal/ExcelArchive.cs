// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.IO.Compression;
using System.Text;

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Thin wrapper around <see cref="ZipArchive"/> that simplifies editing workbook parts in memory.
/// </summary>
internal sealed class ExcelArchive : IDisposable
{
    private readonly MemoryStream _stream;
    private ZipArchive? _zipArchive;
    private bool _disposed;

    private ExcelArchive(byte[] template)
    {
        _stream = new MemoryStream();
        _stream.Write(template, 0, template.Length);
        _stream.Position = 0;
        _zipArchive = new ZipArchive(_stream, ZipArchiveMode.Update, leaveOpen: true);
    }

    /// <summary>
    /// Loads the supplied workbook template bytes into an editable archive.
    /// </summary>
    /// <param name="template">The XLSX template to load.</param>
    /// <returns>An <see cref="ExcelArchive"/> ready for manipulation.</returns>
    public static ExcelArchive Load(byte[] template) => new(template);

    /// <summary>
    /// Serializes the in-memory workbook back into a byte array.
    /// </summary>
    /// <returns>The workbook bytes.</returns>
    public byte[] ToArray()
    {
        _zipArchive?.Dispose();
        _zipArchive = null;
        return _stream.ToArray();
    }

    /// <summary>
    /// Reads the contents of the specified part as UTF-8 text.
    /// </summary>
    /// <param name="path">Part path inside the archive.</param>
    /// <returns>File contents.</returns>
    public string ReadText(string path)
    {
        var entry = GetEntry(path);
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: false);
        return reader.ReadToEnd();
    }
    
    /// <summary>
    /// Writes text into the specified part, truncating existing data.
    /// </summary>
    /// <param name="path">Part path inside the archive.</param>
    /// <param name="content">Text to persist.</param>
    /// <param name="encoding">Optional encoding (defaults to UTF-8 without BOM).</param>
    public void WriteText(string path, string content, Encoding? encoding = null)
    {
        var entry = GetOrCreateEntry(path);
        using var stream = entry.Open();
        stream.SetLength(0);
        using var writer = new StreamWriter(stream, encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: true);
        writer.Write(content);
        writer.Flush();
    }

    /// <summary>
    /// Writes raw bytes into the specified part, truncating existing data.
    /// </summary>
    /// <param name="path">Part path inside the archive.</param>
    /// <param name="content">Data to persist.</param>
    public void WriteBytes(string path, byte[] content)
    {
        var entry = GetOrCreateEntry(path);
        using var stream = entry.Open();
        stream.SetLength(0);
        stream.Write(content, 0, content.Length);
    }

    /// <summary>
    /// Enumerates entries that reside under the provided folder prefix.
    /// </summary>
    /// <param name="folderPrefix">Folder prefix, e.g. <c>xl/tables/</c>.</param>
    /// <returns>Paths of matching entries.</returns>
    public IEnumerable<string> EnumerateEntries(string folderPrefix)
    {
        EnsureNotDisposed();

        foreach (var entry in _zipArchive!.Entries)
        {
            if (entry.FullName.StartsWith(folderPrefix, StringComparison.OrdinalIgnoreCase))
            {
                yield return entry.FullName;
            }
        }
    }

    /// <summary>
    /// Indicates whether the specified entry exists within the archive.
    /// </summary>
    /// <param name="path">Part path inside the archive.</param>
    public bool EntryExists(string path)
    {
        EnsureNotDisposed();
        return _zipArchive!.GetEntry(path) is not null;
    }

    /// <summary>
    /// Removes the specified entry if present.
    /// </summary>
    /// <param name="path">Part path inside the archive.</param>
    public void Remove(string path)
    {
        EnsureNotDisposed();
        _zipArchive!.GetEntry(path)?.Delete();
    }

    private ZipArchiveEntry GetEntry(string path)
    {
        EnsureNotDisposed();
        return _zipArchive?.GetEntry(path) ?? throw new InvalidOperationException($"'{path}' was not found inside the workbook template.");
    }

    private ZipArchiveEntry GetOrCreateEntry(string path)
    {
        EnsureNotDisposed();
        return _zipArchive?.GetEntry(path) ?? _zipArchive!.CreateEntry(path, CompressionLevel.Optimal);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        _zipArchive?.Dispose();
        _zipArchive = null;
        _stream.Dispose();
    }

    private void EnsureNotDisposed()
    {
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ExcelArchive));
        }
    }
}

