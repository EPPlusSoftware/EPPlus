/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

namespace EPPlus.Export.Pdf.Helpers
{
    internal static class BmpDecoder
    {
        private const int BI_RGB = 0;
        private const int BI_RLE8 = 1;
        private const int BI_RLE4 = 2;
        private const int FileHeaderSize = 14;
        private const int MaxDimension = 30000;

        internal static bool IsBmp(byte[] d) => d != null && d.Length >= FileHeaderSize + 40 && d[0] == (byte)'B' && d[1] == (byte)'M';

        internal static bool CanDecode(byte[] d) => TryReadHeader(d, out _);

        internal static bool TryDecode(byte[] d, out int width, out int height, out byte[] rgb)
        {
            width = height = 0;
            rgb = null;
            if (!TryReadHeader(d, out var h))
                return false;

            width = h.Width;
            height = h.Height;
            var outRgb = new byte[width * height * 3];
            if (h.Compression == BI_RLE8 || h.Compression == BI_RLE4)
            {
                DecodeRle(d, h, outRgb);
            }
            else
            {
                switch (h.BitCount)
                {
                    case 24:
                        DecodeTrueColor(d, h, 3, outRgb);
                        break;
                    case 32:
                        DecodeTrueColor(d, h, 4, outRgb);
                        break;
                    case 8:
                        DecodeIndexed(d, h, outRgb);
                        break;
                    case 4:
                        DecodeIndexed(d, h, outRgb);
                        break;
                    case 1:
                        DecodeIndexed(d, h, outRgb);
                        break;
                    default:
                        return false;
                }
            }
            rgb = outRgb;
            return true;
        }

        private struct Header
        {
            public int Width;
            public int Height;
            public bool TopDown;
            public int BitCount;
            public int Compression;
            public int PixelOffset;
            public int RowSize;
            public int PaletteOffset;
            public int PaletteCount;
        }

        private static bool TryReadHeader(byte[] d, out Header h)
        {
            h = default;
            if (!IsBmp(d)) return false;

            int pixelOffset = ReadLE32(d, 10);
            int dibSize = ReadLE32(d, FileHeaderSize);
            if (dibSize < 40 || FileHeaderSize + dibSize > d.Length)
                return false;

            int width = ReadLE32(d, FileHeaderSize + 4);
            int rawHeight = ReadLE32(d, FileHeaderSize + 8);
            int bitCount = ReadLE16(d, FileHeaderSize + 14);
            int compression = ReadLE32(d, FileHeaderSize + 16);
            int colorsUsed = ReadLE32(d, FileHeaderSize + 32);
            bool rle = (compression == BI_RLE8 && bitCount == 8) || (compression == BI_RLE4 && bitCount == 4);
            if (compression != BI_RGB && !rle)
                return false;
            if (width <= 0 || rawHeight == 0)
                return false;

            bool topDown = rawHeight < 0;
            int height = rawHeight < 0 ? -rawHeight : rawHeight;
            if (width > MaxDimension || height > MaxDimension)
                return false;
            if (bitCount != 1 && bitCount != 4 && bitCount != 8 && bitCount != 24 && bitCount != 32)
                return false;
            if (rle && topDown)
                return false;

            if (pixelOffset <= 0) pixelOffset = FileHeaderSize + dibSize;
            int rowSize = 0;
            if (!rle)
            {
                rowSize = ((bitCount * width + 31) / 32) * 4;
                long pixelSpan = (long)rowSize * height;
                if (pixelOffset < 0 || pixelOffset + pixelSpan > d.Length)
                    return false;
            }
            else if (pixelOffset < 0 || pixelOffset >= d.Length)
            {
                return false;
            }

            int paletteOffset = FileHeaderSize + dibSize;
            int paletteCount = 0;
            if (bitCount <= 8)
            {
                paletteCount = colorsUsed > 0 ? colorsUsed : (1 << bitCount);
                if (paletteCount < 0 || paletteCount > 256)
                    return false;
                if (paletteOffset + paletteCount * 4 > d.Length)
                    return false;
            }
            h = new Header
            {
                Width = width,
                Height = height,
                TopDown = topDown,
                BitCount = bitCount,
                Compression = compression,
                PixelOffset = pixelOffset,
                RowSize = rowSize,
                PaletteOffset = paletteOffset,
                PaletteCount = paletteCount,
            };
            return true;
        }

        private static void DecodeTrueColor(byte[] d, Header h, int bytesPerPixel, byte[] outRgb)
        {
            for (int y = 0; y < h.Height; y++)
            {
                int srcRow = h.TopDown ? y : (h.Height - 1 - y);
                int src = h.PixelOffset + srcRow * h.RowSize;
                int dst = y * h.Width * 3;
                for (int x = 0; x < h.Width; x++)
                {
                    int p = src + x * bytesPerPixel;
                    outRgb[dst++] = d[p + 2];
                    outRgb[dst++] = d[p + 1];
                    outRgb[dst++] = d[p];
                }
            }
        }

        private static void DecodeIndexed(byte[] d, Header h, byte[] outRgb)
        {
            for (int y = 0; y < h.Height; y++)
            {
                int srcRow = h.TopDown ? y : (h.Height - 1 - y);
                int src = h.PixelOffset + srcRow * h.RowSize;
                int dst = y * h.Width * 3;
                for (int x = 0; x < h.Width; x++)
                {
                    int index = SampleIndex(d, src, x, h.BitCount);
                    if (index >= h.PaletteCount) index = 0;
                    int pal = h.PaletteOffset + index * 4;
                    outRgb[dst++] = d[pal + 2];
                    outRgb[dst++] = d[pal + 1];
                    outRgb[dst++] = d[pal];
                }
            }
        }

        private static int SampleIndex(byte[] d, int rowStart, int x, int bitCount)
        {
            switch (bitCount)
            {
                case 8:
                    return d[rowStart + x];
                case 4:
                    {
                        byte b = d[rowStart + (x >> 1)];
                        return (x & 1) == 0 ? (b >> 4) : (b & 0x0F);
                    }
                case 1:
                    {
                        byte b = d[rowStart + (x >> 3)];
                        int shift = 7 - (x & 7);
                        return (b >> shift) & 1;
                    }
                default:
                    return 0;
            }
        }

        private static void DecodeRle(byte[] d, Header h, byte[] outRgb)
        {
            int w = h.Width, height = h.Height;
            bool rle4 = h.Compression == BI_RLE4;
            var index = new byte[w * height];
            int pos = h.PixelOffset;
            int x = 0, row = 0;
            while (pos + 1 < d.Length)
            {
                int count = d[pos++];
                int val = d[pos++];
                if (count > 0)
                {
                    for (int i = 0; i < count; i++)
                    {
                        if (row < height && x < w)
                        {
                            int idx = rle4 ? ((i & 1) == 0 ? (val >> 4) : (val & 0x0F)) : val;
                            index[row * w + x] = (byte)idx;
                        }
                        x++;
                    }
                }
                else if (val == 0)
                {
                    x = 0;
                    row++;
                }
                else if (val == 1)
                {
                    break;
                }
                else if (val == 2)
                {
                    if (pos + 1 >= d.Length) break;
                    x += d[pos++];
                    row += d[pos++];
                }
                else
                {
                    int n = val;
                    if (!rle4)
                    {
                        for (int i = 0; i < n; i++)
                        {
                            if (pos >= d.Length) break;
                            int idx = d[pos++];
                            if (row < height && x < w) index[row * w + x] = (byte)idx;
                            x++;
                        }
                        if ((n & 1) != 0) pos++;
                    }
                    else
                    {
                        int bytesNeeded = (n + 1) / 2;
                        for (int i = 0; i < n; i++)
                        {
                            int bp = pos + (i >> 1);
                            if (bp >= d.Length) break;
                            int packed = d[bp];
                            int idx = (i & 1) == 0 ? (packed >> 4) : (packed & 0x0F);
                            if (row < height && x < w) index[row * w + x] = (byte)idx;
                            x++;
                        }
                        pos += bytesNeeded + (bytesNeeded & 1);
                    }
                }
            }
            for (int y = 0; y < height; y++)
            {
                int srcRow = height - 1 - y;
                int dst = y * w * 3;
                for (int xx = 0; xx < w; xx++)
                {
                    int idx = index[srcRow * w + xx];
                    if (idx >= h.PaletteCount) idx = 0;
                    int pal = h.PaletteOffset + idx * 4;
                    outRgb[dst++] = d[pal + 2];
                    outRgb[dst++] = d[pal + 1];
                    outRgb[dst++] = d[pal];
                }
            }
        }

        private static int ReadLE16(byte[] d, int i) => d[i] | (d[i + 1] << 8);
        private static int ReadLE32(byte[] d, int i) => d[i] | (d[i + 1] << 8) | (d[i + 2] << 16) | (d[i + 3] << 24);
    }
}