// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

using System.Text;
using System.Linq;

namespace Microsoft.ConnectedWorkbooks.Internal;

internal static class XmlEncodingHelper
{
    public static string DecodeToString(byte[] xmlBytes)
    {
        if (xmlBytes.Length >= 3 && xmlBytes[0] == 0xEF && xmlBytes[1] == 0xBB && xmlBytes[2] == 0xBF)
        {
            return Encoding.UTF8.GetString(xmlBytes, 3, xmlBytes.Length - 3);
        }

        if (xmlBytes.Length >= 2 && xmlBytes[0] == 0xFF && xmlBytes[1] == 0xFE)
        {
            return Encoding.Unicode.GetString(xmlBytes, 2, xmlBytes.Length - 2);
        }

        if (xmlBytes.Length >= 2 && xmlBytes[0] == 0xFE && xmlBytes[1] == 0xFF)
        {
            return Encoding.BigEndianUnicode.GetString(xmlBytes, 2, xmlBytes.Length - 2);
        }

        return Encoding.UTF8.GetString(xmlBytes);
    }

    public static byte[] EncodeWithBom(string content, Encoding encoding)
    {
        if (encoding == Encoding.Unicode)
        {
            return Encoding.Unicode.GetPreamble().Concat(encoding.GetBytes(content)).ToArray();
        }

        if (encoding == Encoding.BigEndianUnicode)
        {
            return Encoding.BigEndianUnicode.GetPreamble().Concat(encoding.GetBytes(content)).ToArray();
        }

        if (encoding == Encoding.UTF8)
        {
            return Encoding.UTF8.GetPreamble().Concat(encoding.GetBytes(content)).ToArray();
        }

        return encoding.GetBytes(content);
    }
}

