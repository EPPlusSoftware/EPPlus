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

namespace EPPlus.Fonts.OpenType
{
    internal class FontTableReaderFactory
    {
        public FontTableReaderFactory(byte[] fontBytes)
        {
            _fontBytes = fontBytes ?? throw new ArgumentNullException(nameof(fontBytes));
        }

        private readonly byte[] _fontBytes;

        public int FontBytesLength => _fontBytes?.Length ?? 0;

        public FontsBinaryReader CreateReader(long baseOffset = 0)
        {
            if (baseOffset == -1) baseOffset = 0;
            // Validate and clamp baseOffset
            if (baseOffset < 0 || baseOffset > _fontBytes.Length)
                throw new ArgumentOutOfRangeException("baseOffset", "Offset is outside of the font buffer.");

            var ms = new MemoryStream(_fontBytes);
            ms.Position = baseOffset;
            return new FontsBinaryReader(ms);
        }
    }
}
