/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/04/2024         EPPlus Software AB       EPPlus 7.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal static class ByteArrayExtensions
    {
        private static Encoding GetFileEncoding(byte[] buffer)
        {
            // Use the default encoding (usually ANSI)
            Encoding enc = Encoding.UTF8;

            // Detect the encoding based on the initial bytes
            if (buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
                enc = Encoding.UTF8;
            else if (buffer[0] == 0xFF && buffer[1] == 0xFE)
                enc = Encoding.Unicode;
            else if (buffer[0] == 0xFE && buffer[1] == 0xFF)
                enc = Encoding.Unicode; // UTF-16 LE
            else if (buffer[0] == 0xFF && buffer[1] == 0xFE)
                enc = Encoding.BigEndianUnicode; // UTF-16 BE
            else if (buffer[0] == 0 && buffer[1] == 0 && buffer[2] == 0xFE && buffer[3] == 0xFF)
                enc = Encoding.UTF32;

            return enc;
        }

        public static string GetEncodedString(this byte[] bytes, out Encoding enc)
        {
            enc = Encoding.UTF8;
            if (bytes == null || bytes.Length < 4) return null;
            enc = GetFileEncoding(bytes);
            return enc.GetString(bytes);
        }
    }
}
