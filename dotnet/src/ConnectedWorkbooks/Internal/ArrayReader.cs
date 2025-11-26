// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Buffers.Binary;

namespace Microsoft.ConnectedWorkbooks.Internal;

/// <summary>
/// Lightweight helper for sequentially reading primitive values from a byte buffer.
/// </summary>
internal sealed class ArrayReader
{
    private readonly ReadOnlyMemory<byte> _buffer;
    private int _offset;

    /// <summary>
    /// Initializes a new reader that iterates over the supplied byte array.
    /// </summary>
    /// <param name="buffer">The byte array to consume.</param>
    public ArrayReader(byte[] buffer)
        : this(new ReadOnlyMemory<byte>(buffer))
    {
    }

    /// <summary>
    /// Initializes a new reader for the provided byte memory segment.
    /// </summary>
    /// <param name="buffer">The data segment to consume.</param>
    public ArrayReader(ReadOnlyMemory<byte> buffer)
    {
        _buffer = buffer;
        _offset = 0;
    }

    /// <summary>
    /// Reads the specified number of bytes, advancing the reader past them.
    /// </summary>
    /// <param name="count">Number of bytes to read.</param>
    /// <returns>The requested slice of the underlying buffer.</returns>
    public ReadOnlyMemory<byte> ReadMemory(int count)
    {
        EnsureAvailable(count);
        var slice = _buffer.Slice(_offset, count);
        _offset += count;
        return slice;
    }

    /// <summary>
    /// Reads a 32-bit little-endian integer from the buffer.
    /// </summary>
    /// <returns>The parsed integer.</returns>
    public int ReadInt32()
    {
        EnsureAvailable(sizeof(int));
        var value = BinaryPrimitives.ReadInt32LittleEndian(_buffer.Span.Slice(_offset, sizeof(int)));
        _offset += sizeof(int);
        return value;
    }

    /// <summary>
    /// Returns the remaining bytes from the current position and advances to the end.
    /// </summary>
    /// <returns>A slice containing all remaining bytes.</returns>
    public ReadOnlyMemory<byte> ReadToEnd()
    {
        var slice = _buffer.Slice(_offset);
        _offset = _buffer.Length;
        return slice;
    }

    private void EnsureAvailable(int count)
    {
        if (_offset + count > _buffer.Length)
        {
            throw new InvalidOperationException("Attempted to read beyond the length of the buffer.");
        }
    }
}

