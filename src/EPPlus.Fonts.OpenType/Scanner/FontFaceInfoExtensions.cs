/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using System;
using System.IO;

// File: FontFaceInfoExtensions.cs
namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Helper extensions to make tests as clean as before.
    /// </summary>
    public static class FontFaceInfoExtensions
    {
        /// <summary>
        /// Returns the raw bytes of a font table – exactly like the old ScannedFont.GetTableBytes().
        /// Thread-safe, fast, and zero memory leaks.
        /// </summary>
        public static byte[] GetTableBytes(this FontFaceInfo face, string tag)
        {
            if (face == null) throw new ArgumentNullException(nameof(face));
            if (!face.TableRecords.TryGetValue(tag, out var record))
                throw new ArgumentException($"Table '{tag}' not found in font {face.FilePath}");

            using var fs = new FileStream(face.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read);
            fs.Position = face.OffsetInFile + record.Offset;

            var buffer = new byte[record.Length];
            int read = fs.Read(buffer, 0, buffer.Length);
            if (read != buffer.Length)
                throw new IOException($"Could not read entire table '{tag}' from {face.FilePath}");

            return buffer;
        }
    }
}