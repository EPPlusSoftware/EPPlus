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
