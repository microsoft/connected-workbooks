// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Buffers.Binary;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal sealed class ArrayReader
{
    private readonly ReadOnlyMemory<byte> _buffer;
    private int _offset;

    public ArrayReader(byte[] buffer)
        : this(new ReadOnlyMemory<byte>(buffer))
    {
    }

    public ArrayReader(ReadOnlyMemory<byte> buffer)
    {
        _buffer = buffer;
        _offset = 0;
    }

    public ReadOnlyMemory<byte> ReadMemory(int count)
    {
        EnsureAvailable(count);
        var slice = _buffer.Slice(_offset, count);
        _offset += count;
        return slice;
    }

    public int ReadInt32()
    {
        EnsureAvailable(sizeof(int));
        var value = BinaryPrimitives.ReadInt32LittleEndian(_buffer.Span.Slice(_offset, sizeof(int)));
        _offset += sizeof(int);
        return value;
    }

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

