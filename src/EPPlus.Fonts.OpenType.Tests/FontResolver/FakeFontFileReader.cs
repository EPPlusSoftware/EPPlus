/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/06/2026         EPPlus Software AB           Initial implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Scanner;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.FontResolver
{
    /// <summary>
    /// Test fake of IFontFileReader. Returns predefined bytes for registered file paths.
    /// An unregistered path is a test bug — the resolver tried to read a face the scanner
    /// returned but the test forgot to register bytes for it — so we throw.
    /// </summary>
    internal sealed class FakeFontFileReader : IFontFileReader
    {
        private readonly Dictionary<string, byte[]> _bytesByPath =
            new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Registers explicit bytes for the given file path. Returns this for fluent chaining.
        /// </summary>
        public FakeFontFileReader Register(string filePath, byte[] bytes)
        {
            _bytesByPath[filePath] = bytes;
            return this;
        }

        /// <summary>
        /// Convenience overload: encodes a marker string as UTF-8 bytes and registers them.
        /// Useful when the test only cares about identifying which fake was returned, not
        /// the exact contents.
        /// </summary>
        public FakeFontFileReader Register(string filePath, string marker)
        {
            return Register(filePath, Encoding.UTF8.GetBytes(marker));
        }

        public byte[] ReadFontBytes(FontFaceInfo face)
        {
            if (face == null)
                throw new ArgumentNullException("face");

            if (_bytesByPath.TryGetValue(face.FilePath, out var bytes))
                return bytes;

            throw new InvalidOperationException(
                "FakeFontFileReader has no bytes registered for path: " + face.FilePath +
                ". Test arrange is incomplete — register bytes for every path the FakeFontScanner returns.");
        }
    }
}