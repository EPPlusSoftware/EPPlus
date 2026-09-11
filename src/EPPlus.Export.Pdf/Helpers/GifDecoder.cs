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
using System.IO;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal static class GifDecoder
    {
        private const int MaxDimension = 30000;

        internal static bool IsGif(byte[] d)  => d != null && d.Length > 13 && d[0] == 'G' && d[1] == 'I' && d[2] == 'F' && d[3] == '8' && (d[4] == '7' || d[4] == '9') && d[5] == 'a';

        internal static bool CanDecode(byte[] d) => TryDecode(d, out _, out _, out _, out _);

        internal static bool HasTransparency(byte[] d)
        {
            if (!IsGif(d)) return false;
            if (!TryReadScreen(d, out _, out _, out int pos, out _, out _)) return false;
            return ScanToFirstImage(d, ref pos, out int transparentIndex) && transparentIndex >= 0;
        }

        internal static bool TryDecode(byte[] d, out int width, out int height, out byte[] rgb, out byte[] alpha)
        {
            width = height = 0;
            rgb = null;
            alpha = null;
            if (!TryReadScreen(d, out int screenW, out int screenH, out int pos, out byte[] gct, out int _))
                return false;
            if (screenW <= 0 || screenH <= 0 || screenW > MaxDimension || screenH > MaxDimension)
                return false;
            if (!ScanToFirstImage(d, ref pos, out int transparentIndex))
                return false;

            if (pos + 10 > d.Length || d[pos] != 0x2C) return false;
            int left = ReadLE16(d, pos + 1);
            int top = ReadLE16(d, pos + 3);
            int fw = ReadLE16(d, pos + 5);
            int fh = ReadLE16(d, pos + 7);
            int packed = d[pos + 9];
            pos += 10;
            if (fw <= 0 || fh <= 0 || fw > MaxDimension || fh > MaxDimension)
                return false;

            bool lctFlag = (packed & 0x80) != 0;
            bool interlace = (packed & 0x40) != 0;
            byte[] palette = gct;
            if (lctFlag)
            {
                int lctSize = 2 << (packed & 7);
                if (pos + lctSize * 3 > d.Length)
                    return false;
                palette = new byte[lctSize * 3];
                System.Array.Copy(d, pos, palette, 0, lctSize * 3);
                pos += lctSize * 3;
            }
            if (palette == null)
                return false;

            int paletteCount = palette.Length / 3;
            if (pos >= d.Length)
                return false;

            int minCodeSize = d[pos++];
            if (minCodeSize < 2 || minCodeSize > 8)
                return false;

            byte[] lzw = ReadSubBlocks(d, ref pos);
            var indices = new byte[fw * fh];
            if (!LzwDecode(lzw, minCodeSize, indices))
                return false;

            if (interlace) indices = Deinterlace(indices, fw, fh);
            bool hasAlpha = transparentIndex >= 0;
            rgb = new byte[screenW * screenH * 3];
            if (hasAlpha) alpha = new byte[screenW * screenH];

            for (int fy = 0; fy < fh; fy++)
            {
                int cy = top + fy;
                if (cy < 0 || cy >= screenH) continue;
                for (int fx = 0; fx < fw; fx++)
                {
                    int cx = left + fx;
                    if (cx < 0 || cx >= screenW) continue;
                    int index = indices[fy * fw + fx];
                    int canvas = cy * screenW + cx;
                    if (hasAlpha && index == transparentIndex)
                        continue;
                    int p = index < paletteCount ? index * 3 : 0;
                    rgb[canvas * 3] = palette[p];
                    rgb[canvas * 3 + 1] = palette[p + 1];
                    rgb[canvas * 3 + 2] = palette[p + 2];
                    if (hasAlpha) alpha[canvas] = 255;
                }
            }
            width = screenW;
            height = screenH;
            return true;
        }

        private static bool TryReadScreen(byte[] d, out int width, out int height, out int pos, out byte[] gct, out int bgIndex)
        {
            width = height = 0; pos = 0; gct = null; bgIndex = 0;
            if (!IsGif(d))
                return false;
            width = ReadLE16(d, 6);
            height = ReadLE16(d, 8);
            int packed = d[10];
            bgIndex = d[11];
            pos = 13;
            if ((packed & 0x80) != 0)
            {
                int gctSize = 2 << (packed & 7);
                if (pos + gctSize * 3 > d.Length) return false;
                gct = new byte[gctSize * 3];
                System.Array.Copy(d, pos, gct, 0, gctSize * 3);
                pos += gctSize * 3;
            }
            return true;
        }

        private static bool ScanToFirstImage(byte[] d, ref int pos, out int transparentIndex)
        {
            transparentIndex = -1;
            while (pos < d.Length)
            {
                int b = d[pos];
                if (b == 0x2C) return true;
                if (b == 0x3B) return false;
                if (b == 0x21)
                {
                    if (pos + 2 > d.Length)
                        return false;

                    int label = d[pos + 1];
                    pos += 2;
                    if (label == 0xF9)
                    {
                        if (pos >= d.Length) return false;
                        int size = d[pos];
                        if (size >= 4 && pos + 1 + size <= d.Length)
                        {
                            int gcePacked = d[pos + 1];
                            if ((gcePacked & 0x01) != 0) transparentIndex = d[pos + 4];
                        }
                        SkipSubBlocks(d, ref pos);
                    }
                    else
                    {
                        SkipSubBlocks(d, ref pos);
                    }
                }
                else
                {
                    return false;
                }
            }
            return false;
        }

        private static void SkipSubBlocks(byte[] d, ref int pos)
        {
            while (pos < d.Length)
            {
                int len = d[pos++];
                if (len == 0) break;
                pos += len;
            }
        }

        private static byte[] ReadSubBlocks(byte[] d, ref int pos)
        {
            using (var ms = new MemoryStream())
            {
                while (pos < d.Length)
                {
                    int len = d[pos++];
                    if (len == 0) break;
                    if (pos + len > d.Length) { len = d.Length - pos; ms.Write(d, pos, len); pos += len; break; }
                    ms.Write(d, pos, len);
                    pos += len;
                }
                return ms.ToArray();
            }
        }

        private static bool LzwDecode(byte[] data, int minCodeSize, byte[] outIndices)
        {
            const int MaxCodes = 4096;
            int clearCode = 1 << minCodeSize;
            int endCode = clearCode + 1;
            var prefix = new int[MaxCodes];
            var suffix = new byte[MaxCodes];
            for (int i = 0; i < clearCode; i++)
            {
                prefix[i] = -1; suffix[i] = (byte)i;
            }

            int codeSize = minCodeSize + 1;
            int nextCode = endCode + 1;
            int prev = -1;
            int outPos = 0;
            int bitPos = 0;
            var stack = new byte[MaxCodes + 1];

            while (true)
            {
                int code = ReadCode(data, ref bitPos, codeSize);
                if (code < 0)
                    break;
                if (code == clearCode)
                {
                    codeSize = minCodeSize + 1;
                    nextCode = endCode + 1;
                    prev = -1;
                    continue;
                }
                if (code == endCode)
                    break;
                if (prev < 0)
                {
                    if (code >= clearCode)
                        return false;
                    if (outPos < outIndices.Length) outIndices[outPos++] = (byte)code;
                    prev = code;
                    continue;
                }
                int emit;
                bool kwk = false;
                if (code < nextCode)
                {
                    emit = code;
                }
                else if (code == nextCode)
                {
                    emit = prev;
                    kwk = true;
                }
                else
                {
                    return false;
                }
                int sp = 0;
                int c = emit;
                while (c >= 0)
                {
                    if (sp >= stack.Length)
                        return false;
                    stack[sp++] = suffix[c];
                    c = prefix[c];
                }
                byte firstByte = stack[sp - 1];
                for (int i = sp - 1; i >= 0 && outPos < outIndices.Length; i--)
                    outIndices[outPos++] = stack[i];
                if (kwk && outPos < outIndices.Length)
                    outIndices[outPos++] = firstByte;
                if (nextCode < MaxCodes)
                {
                    prefix[nextCode] = prev;
                    suffix[nextCode] = firstByte;
                    nextCode++;
                    if (nextCode == (1 << codeSize) && codeSize < 12) codeSize++;
                }
                prev = code;
                if (outPos >= outIndices.Length)
                    break;
            }
            return true;
        }

        private static int ReadCode(byte[] data, ref int bitPos, int codeSize)
        {
            int code = 0;
            for (int i = 0; i < codeSize; i++)
            {
                int bytePos = bitPos >> 3;
                if (bytePos >= data.Length)
                    return -1;
                int bit = (data[bytePos] >> (bitPos & 7)) & 1;
                code |= bit << i;
                bitPos++;
            }
            return code;
        }

        private static byte[] Deinterlace(byte[] src, int w, int h)
        {
            var dst = new byte[w * h];
            int[] starts = { 0, 4, 2, 1 };
            int[] steps = { 8, 8, 4, 2 };
            int srcRow = 0;
            for (int pass = 0; pass < 4; pass++)
            {
                for (int y = starts[pass]; y < h; y += steps[pass])
                {
                    System.Array.Copy(src, srcRow * w, dst, y * w, w);
                    srcRow++;
                }
            }
            return dst;
        }

        private static int ReadLE16(byte[] d, int i) => d[i] | (d[i + 1] << 8);
    }
}