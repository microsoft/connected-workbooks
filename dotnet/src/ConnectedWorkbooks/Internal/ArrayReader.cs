// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Buffers.Binary;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal sealed class ArrayReader
{
    private readonly ReadOnlyMemory<byte> _buffer;
    private int _offset;

    public ArrayReader(byte[] buffer)
    {
        _buffer = new ReadOnlyMemory<byte>(buffer);
        _offset = 0;
    }

    public byte[] ReadBytes(int count)
    {
        if (_offset + count > _buffer.Length)
        {
            throw new InvalidOperationException("Attempted to read beyond the length of the buffer.");
        }

        var slice = _buffer.Slice(_offset, count).ToArray();
        _offset += count;
        return slice;
    }

    public int ReadInt32()
    {
        var span = _buffer.Span;
        if (_offset + sizeof(int) > span.Length)
        {
            throw new InvalidOperationException("Attempted to read beyond the length of the buffer.");
        }

        var value = BinaryPrimitives.ReadInt32LittleEndian(span.Slice(_offset, sizeof(int)));
        _offset += sizeof(int);
        return value;
    }

    public byte[] ReadToEnd()
    {
        var slice = _buffer.Slice(_offset).ToArray();
        _offset = _buffer.Length;
        return slice;
    }
}

