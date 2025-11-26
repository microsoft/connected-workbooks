// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.IO.Compression;
using System.Text;

namespace Microsoft.ConnectedWorkbooks.Internal;

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

    public static ExcelArchive Load(byte[] template) => new(template);

    public byte[] ToArray()
    {
        _zipArchive?.Dispose();
        _zipArchive = null;
        return _stream.ToArray();
    }

    public string ReadText(string path)
    {
        var entry = GetEntry(path);
        using var stream = entry.Open();
        using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, leaveOpen: false);
        return reader.ReadToEnd();
    }

    public byte[] ReadBytes(string path)
    {
        var entry = GetEntry(path);
        using var stream = entry.Open();
        using var memory = new MemoryStream();
        stream.CopyTo(memory);
        return memory.ToArray();
    }

    public void WriteText(string path, string content, Encoding? encoding = null)
    {
        var entry = GetOrCreateEntry(path);
        using var stream = entry.Open();
        stream.SetLength(0);
        using var writer = new StreamWriter(stream, encoding ?? new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), leaveOpen: true);
        writer.Write(content);
        writer.Flush();
    }

    public void WriteBytes(string path, byte[] content)
    {
        var entry = GetOrCreateEntry(path);
        using var stream = entry.Open();
        stream.SetLength(0);
        stream.Write(content, 0, content.Length);
    }

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

    public bool EntryExists(string path)
    {
        EnsureNotDisposed();
        return _zipArchive!.GetEntry(path) is not null;
    }

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

