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

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal static class IcoDecoder
    {
        private const int MaxDimension = 30000;
        private const int DirEntrySize = 16;

        internal static bool IsIco(byte[] d)
        {
            if (d == null || d.Length < 6 + DirEntrySize) return false;
            if (ReadLE16(d, 0) != 0 || ReadLE16(d, 2) != 1) return false;
            int count = ReadLE16(d, 4);
            return count >= 1 && 6 + count * DirEntrySize <= d.Length;
        }

        internal static bool TryGetBestFrame(byte[] d, out byte[] frame, out bool selfContained)
        {
            frame = null;
            selfContained = false;
            if (!IsIco(d)) return false;
            int count = ReadLE16(d, 4);

            int bestArea = -1, bestBits = -1, bestOffset = 0, bestSize = 0;
            for (int i = 0; i < count; i++)
            {
                int e = 6 + i * DirEntrySize;
                int w = d[e]; if (w == 0) w = 256;
                int h = d[e + 1]; if (h == 0) h = 256;
                int bits = ReadLE16(d, e + 6);
                int size = ReadLE32(d, e + 8);
                int offset = ReadLE32(d, e + 12);
                if (size <= 0 || offset < 0 || (long)offset + size > d.Length) continue;
                int area = w * h;
                if (area > bestArea || (area == bestArea && bits > bestBits))
                {
                    bestArea = area; bestBits = bits; bestOffset = offset; bestSize = size;
                }
            }
            if (bestArea < 0) return false;

            frame = new byte[bestSize];
            System.Array.Copy(d, bestOffset, frame, 0, bestSize);
            selfContained = IsPngFrame(frame) || IsJpegFrame(frame);
            return true;
        }

        internal static bool CanDecodeDib(byte[] d) => TryDecodeDib(d, out _, out _, out _, out _);

        internal static bool DibHasTransparency(byte[] d) => TryDecodeDib(d, out _, out _, out _, out byte[] a) && a != null;

        internal static bool TryDecodeDib(byte[] d, out int width, out int height, out byte[] rgb, out byte[] alpha)
        {
            width = height = 0;
            rgb = null;
            alpha = null;
            if (d == null || d.Length < 40) return false;

            int dibSize = ReadLE32(d, 0);
            if (dibSize < 40 || dibSize > 124 || dibSize > d.Length) return false;
            int w = ReadLE32(d, 4);
            int doubledHeight = ReadLE32(d, 8);
            int bitCount = ReadLE16(d, 14);
            int compression = ReadLE32(d, 16);
            int colorsUsed = ReadLE32(d, 32);
            if (compression != 0) return false;
            if (w <= 0 || doubledHeight <= 0 || (doubledHeight & 1) != 0) return false;
            int h = doubledHeight / 2;
            if (w > MaxDimension || h > MaxDimension || h <= 0) return false;
            if (bitCount != 1 && bitCount != 4 && bitCount != 8 && bitCount != 24 && bitCount != 32) return false;

            int paletteOffset = dibSize;
            int paletteCount = 0;
            if (bitCount <= 8)
            {
                paletteCount = colorsUsed > 0 ? colorsUsed : (1 << bitCount);
                if (paletteCount < 0 || paletteCount > 256) return false;
                if (paletteOffset + paletteCount * 4 > d.Length) return false;
            }
            int colorOffset = paletteOffset + paletteCount * 4;
            int colorRow = ((bitCount * w + 31) / 32) * 4;
            long colorSize = (long)colorRow * h;
            if (colorOffset + colorSize > d.Length) return false;

            int andOffset = colorOffset + (int)colorSize;
            int andRow = ((w + 31) / 32) * 4;
            bool hasAnd = andOffset + (long)andRow * h <= d.Length;
            var outRgb = new byte[w * h * 3];
            var outAlpha = new byte[w * h];
            bool anyTransparent = false;
            int colorAlphaMax = 0;
            for (int y = 0; y < h; y++)
            {
                int srcRow = h - 1 - y;
                int cRow = colorOffset + srcRow * colorRow;
                int aRow = andOffset + srcRow * andRow;
                int dst = y * w;
                for (int x = 0; x < w; x++)
                {
                    int r, g, b, ca = 255;
                    if (bitCount == 24 || bitCount == 32)
                    {
                        int bpp = bitCount == 24 ? 3 : 4;
                        int p = cRow + x * bpp;
                        b = d[p]; g = d[p + 1]; r = d[p + 2];
                        if (bitCount == 32) { ca = d[p + 3]; if (ca > colorAlphaMax) colorAlphaMax = ca; }
                    }
                    else
                    {
                        int index = SampleIndex(d, cRow, x, bitCount);
                        if (index >= paletteCount) index = 0;
                        int pal = paletteOffset + index * 4;
                        b = d[pal]; g = d[pal + 1]; r = d[pal + 2];
                    }
                    outRgb[dst * 3] = (byte)r;
                    outRgb[dst * 3 + 1] = (byte)g;
                    outRgb[dst * 3 + 2] = (byte)b;
                    int andBit = 0;
                    if (hasAnd)
                    {
                        byte ab = d[aRow + (x >> 3)];
                        andBit = (ab >> (7 - (x & 7))) & 1;
                    }
                    int a = bitCount == 32 ? ca : 255;
                    if (andBit == 1) a = 0;
                    if (a != 255) anyTransparent = true;
                    outAlpha[dst] = (byte)a;
                    dst++;
                }
            }
            if (bitCount == 32 && colorAlphaMax == 0)
            {
                anyTransparent = false;
                for (int y = 0; y < h; y++)
                {
                    int srcRow = h - 1 - y;
                    int aRow = andOffset + srcRow * andRow;
                    for (int x = 0; x < w; x++)
                    {
                        int andBit = 0;
                        if (hasAnd)
                        {
                            byte ab = d[aRow + (x >> 3)];
                            andBit = (ab >> (7 - (x & 7))) & 1;
                        }
                        int a = andBit == 1 ? 0 : 255;
                        if (a != 255) anyTransparent = true;
                        outAlpha[y * w + x] = (byte)a;
                    }
                }
            }
            width = w;
            height = h;
            rgb = outRgb;
            alpha = anyTransparent ? outAlpha : null;
            return true;
        }

        private static int SampleIndex(byte[] d, int rowStart, int x, int bitCount)
        {
            switch (bitCount)
            {
                case 8:
                    return d[rowStart + x];
                case 4:
                    {
                        byte bb = d[rowStart + (x >> 1)];
                        return (x & 1) == 0 ? (bb >> 4) : (bb & 0x0F);
                    }
                case 1:
                    {
                        byte bb = d[rowStart + (x >> 3)];
                        return (bb >> (7 - (x & 7))) & 1;
                    }
                default:
                    return 0;
            }
        }

        private static readonly byte[] _pngSig = { 137, 80, 78, 71, 13, 10, 26, 10 };
        private static bool IsPngFrame(byte[] d)
        {
            if (d == null || d.Length < _pngSig.Length) return false;
            for (int i = 0; i < _pngSig.Length; i++) if (d[i] != _pngSig[i]) return false;
            return true;
        }
        private static bool IsJpegFrame(byte[] d) => d != null && d.Length > 2 && d[0] == 0xFF && d[1] == 0xD8;

        private static int ReadLE16(byte[] d, int i) => d[i] | (d[i + 1] << 8);
        private static int ReadLE32(byte[] d, int i) => d[i] | (d[i + 1] << 8) | (d[i + 2] << 16) | (d[i + 3] << 24);
    }
}