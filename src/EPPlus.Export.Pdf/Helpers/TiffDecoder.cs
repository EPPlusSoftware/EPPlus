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
using OfficeOpenXml.Packaging.Ionic.Zlib;
using System.IO;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal static class TiffDecoder
    {
        private const int MaxDimension = 30000;

        internal static bool IsTiff(byte[] d)
            => d != null && d.Length >= 8 &&
               ((d[0] == 'I' && d[1] == 'I' && d[2] == 42 && d[3] == 0) ||
                (d[0] == 'M' && d[1] == 'M' && d[2] == 0 && d[3] == 42));

        // TIFF has too many variants to gate cheaply, and images in a sheet are few, so CanEmbed just
        // attempts the decode.
        internal static bool CanDecode(byte[] d) => TryDecode(d, out _, out _, out _, out _);

        // For ProducesSoftMask: does the image carry (non-opaque) alpha?
        internal static bool HasTransparency(byte[] d)
            => TryDecode(d, out _, out _, out _, out byte[] a) && a != null;

        internal static bool TryDecode(byte[] d, out int width, out int height, out byte[] rgb, out byte[] alpha)
        {
            width = height = 0;
            rgb = null;
            alpha = null;
            if (!ReadHeader(d, out bool be, out int ifd)) return false;
            if (ifd + 2 > d.Length) return false;
            int entries = ReadU16(d, be, ifd);
            if (entries <= 0 || (long)ifd + 2 + entries * 12 + 4 > d.Length) return false;

            long imgW = 0, imgH = 0, compression = 1, photometric = -1, spp = 1, rowsPerStrip = 0, planar = 1, predictor = 1;
            long[] bitsPerSample = null, stripOffsets = null, stripByteCounts = null, colorMap = null, extraSamples = null;
            bool tiled = false;

            for (int i = 0; i < entries; i++)
            {
                int e = ifd + 2 + i * 12;
                int tag = ReadU16(d, be, e);
                int type = ReadU16(d, be, e + 2);
                int cnt = (int)ReadU32(d, be, e + 4);
                switch (tag)
                {
                    case 256: imgW = ReadValue(d, be, type, e + 8); break;             // ImageWidth
                    case 257: imgH = ReadValue(d, be, type, e + 8); break;             // ImageLength
                    case 258: bitsPerSample = ReadValues(d, be, type, cnt, e + 8); break;
                    case 259: compression = ReadValue(d, be, type, e + 8); break;
                    case 262: photometric = ReadValue(d, be, type, e + 8); break;
                    case 273: stripOffsets = ReadValues(d, be, type, cnt, e + 8); break;
                    case 277: spp = ReadValue(d, be, type, e + 8); break;              // SamplesPerPixel
                    case 278: rowsPerStrip = ReadValue(d, be, type, e + 8); break;
                    case 279: stripByteCounts = ReadValues(d, be, type, cnt, e + 8); break;
                    case 284: planar = ReadValue(d, be, type, e + 8); break;
                    case 317: predictor = ReadValue(d, be, type, e + 8); break;
                    case 320: colorMap = ReadValues(d, be, type, cnt, e + 8); break;
                    case 322: case 323: case 324: case 325: tiled = true; break;       // tile tags
                    case 338: extraSamples = ReadValues(d, be, type, cnt, e + 8); break;
                }
            }

            if (imgW <= 0 || imgH <= 0 || imgW > MaxDimension || imgH > MaxDimension) return false;
            if (tiled || planar != 1) return false;
            if (stripOffsets == null || stripByteCounts == null) return false;
            if (predictor != 1 && predictor != 2) return false;
            // none / LZW / Deflate (Adobe 8 and 32946) / PackBits.
            if (compression != 1 && compression != 5 && compression != 8 && compression != 32946 && compression != 32773) return false;

            int w = (int)imgW, h = (int)imgH, samples = (int)spp;
            int bits = (bitsPerSample != null && bitsPerSample.Length > 0) ? (int)bitsPerSample[0] : 1;
            if (bitsPerSample != null)
                foreach (var b in bitsPerSample) if (b != bits) return false;         // uniform depth only

            bool isRgb = photometric == 2;
            bool isPalette = photometric == 3;
            bool isGray = photometric == 0 || photometric == 1;
            if (isRgb) { if (bits != 8 || (samples != 3 && samples != 4)) return false; }
            else if (isPalette) { if (bits != 8 || samples != 1 || colorMap == null) return false; }
            else if (isGray) { if (samples != 1 || (bits != 1 && bits != 8)) return false; }
            else return false;

            int rowSize = (w * samples * bits + 7) / 8;
            if (rowsPerStrip <= 0) rowsPerStrip = h;                                    // default: single strip
            int strips = (int)((h + rowsPerStrip - 1) / rowsPerStrip);
            if (stripOffsets.Length < strips || stripByteCounts.Length < strips) return false;

            // Decompress every strip into one contiguous image buffer.
            var image = new byte[rowSize * h];
            int imgPos = 0;
            for (int s = 0; s < strips; s++)
            {
                int rowsInStrip = (int)System.Math.Min(rowsPerStrip, h - (long)s * rowsPerStrip);
                int expected = rowSize * rowsInStrip;
                int off = (int)stripOffsets[s];
                int len = (int)stripByteCounts[s];
                if (off < 0 || len < 0 || (long)off + len > d.Length) return false;
                byte[] raw = Slice(d, off, len);
                byte[] strip;
                switch (compression)
                {
                    case 1: strip = raw; break;
                    case 5: strip = LzwDecode(raw, expected); break;
                    case 32773: strip = PackBits(raw, expected); break;
                    default: strip = Inflate(raw); break;                              // 8 or 32946
                }
                if (strip == null) return false;
                int copy = System.Math.Min(expected, strip.Length);
                if (imgPos + copy > image.Length) copy = image.Length - imgPos;
                System.Array.Copy(strip, 0, image, imgPos, copy);
                imgPos += expected;
            }

            if (predictor == 2 && bits == 8)
                ApplyHorizontalPredictor(image, w, h, samples, rowSize);

            // Build the palette lookup once (16-bit colour map values, or 8-bit if the writer stored them
            // small). Layout is all reds, then all greens, then all blues.
            byte[] palR = null, palG = null, palB = null;
            if (isPalette) BuildPalette(colorMap, out palR, out palG, out palB);

            rgb = new byte[w * h * 3];
            byte[] a = (isRgb && samples == 4) ? new byte[w * h] : null;
            bool premultiplied = extraSamples != null && extraSamples.Length > 0 && extraSamples[0] == 1;
            bool anyAlpha = false;

            for (int y = 0; y < h; y++)
            {
                int row = y * rowSize;
                for (int x = 0; x < w; x++)
                {
                    int di = (y * w + x) * 3;
                    if (isGray)
                    {
                        int gray;
                        if (bits == 8) gray = image[row + x];
                        else gray = ((image[row + (x >> 3)] >> (7 - (x & 7))) & 1) * 255;
                        if (photometric == 0) gray = 255 - gray;                       // WhiteIsZero
                        rgb[di] = rgb[di + 1] = rgb[di + 2] = (byte)gray;
                    }
                    else if (isPalette)
                    {
                        int idx = image[row + x];
                        rgb[di] = palR[idx]; rgb[di + 1] = palG[idx]; rgb[di + 2] = palB[idx];
                    }
                    else // RGB / RGBA, chunky
                    {
                        int si = row + x * samples;
                        int r = image[si], g = image[si + 1], b = image[si + 2];
                        if (samples == 4)
                        {
                            int al = image[si + 3];
                            if (premultiplied && al > 0)                               // straighten premultiplied colour
                            {
                                r = System.Math.Min(255, r * 255 / al);
                                g = System.Math.Min(255, g * 255 / al);
                                b = System.Math.Min(255, b * 255 / al);
                            }
                            a[y * w + x] = (byte)al;
                            if (al != 255) anyAlpha = true;
                        }
                        rgb[di] = (byte)r; rgb[di + 1] = (byte)g; rgb[di + 2] = (byte)b;
                    }
                }
            }

            width = w;
            height = h;
            alpha = anyAlpha ? a : null;                                               // no mask when fully opaque
            return true;
        }

        private static void BuildPalette(long[] colorMap, out byte[] r, out byte[] g, out byte[] b)
        {
            int n = colorMap.Length / 3;
            r = new byte[256]; g = new byte[256]; b = new byte[256];
            // Some writers store 8-bit values in the 16-bit map; if nothing exceeds 255, use them directly.
            long max = 0;
            foreach (var v in colorMap) if (v > max) max = v;
            bool wide = max > 255;
            for (int i = 0; i < n && i < 256; i++)
            {
                r[i] = (byte)(wide ? (colorMap[i] >> 8) : colorMap[i]);
                g[i] = (byte)(wide ? (colorMap[n + i] >> 8) : colorMap[n + i]);
                b[i] = (byte)(wide ? (colorMap[2 * n + i] >> 8) : colorMap[2 * n + i]);
            }
        }

        // Horizontal differencing (predictor 2): each 8-bit sample is stored as its difference from the
        // sample one pixel to the left (per channel), reset at the start of each row.
        private static void ApplyHorizontalPredictor(byte[] image, int w, int h, int spp, int rowSize)
        {
            for (int y = 0; y < h; y++)
            {
                int row = y * rowSize;
                for (int x = 1; x < w; x++)
                    for (int c = 0; c < spp; c++)
                        image[row + x * spp + c] = (byte)(image[row + x * spp + c] + image[row + (x - 1) * spp + c]);
            }
        }

        // PackBits run-length decoding (TIFF 32773).
        private static byte[] PackBits(byte[] s, int expected)
        {
            var outBuf = new byte[expected];
            int o = 0, p = 0;
            while (p < s.Length && o < expected)
            {
                sbyte n = (sbyte)s[p++];
                if (n >= 0)
                {
                    int cnt = n + 1;
                    for (int i = 0; i < cnt && p < s.Length && o < expected; i++) outBuf[o++] = s[p++];
                }
                else if (n != -128)
                {
                    int cnt = 1 - n;
                    if (p < s.Length)
                    {
                        byte v = s[p++];
                        for (int i = 0; i < cnt && o < expected; i++) outBuf[o++] = v;
                    }
                }
                // n == -128 is a no-op
            }
            return outBuf;
        }

        private static byte[] Inflate(byte[] s)
        {
            try
            {
                using (var input = new MemoryStream(s))
                using (var z = new ZlibStream(input, CompressionMode.Decompress))
                using (var output = new MemoryStream())
                {
                    var buf = new byte[8192];
                    int n;
                    while ((n = z.Read(buf, 0, buf.Length)) > 0) output.Write(buf, 0, n);
                    return output.ToArray();
                }
            }
            catch { return null; }
        }

        // TIFF LZW (compression 5): codes are packed most-significant-bit first, and the code width steps
        // up one code early — when the next code reaches 2^width - 1 rather than 2^width.
        private static byte[] LzwDecode(byte[] s, int expected)
        {
            const int Clear = 256, Eoi = 257, MaxCodes = 4096;
            var prefix = new int[MaxCodes];
            var suffix = new byte[MaxCodes];
            for (int i = 0; i < 256; i++) { prefix[i] = -1; suffix[i] = (byte)i; }
            var stack = new byte[MaxCodes + 1];

            var outBuf = new byte[expected];
            int outPos = 0;
            int codeWidth = 9, nextCode = 258, prev = -1;
            int bitPos = 0, endBit = s.Length * 8;

            while (true)
            {
                int code = ReadCodeMsb(s, ref bitPos, codeWidth, endBit);
                if (code < 0 || code == Eoi) break;
                if (code == Clear)
                {
                    codeWidth = 9; nextCode = 258; prev = -1;
                    continue;
                }
                if (prev < 0)
                {
                    if (code >= 256) return null;                       // first code after a clear is a literal
                    if (outPos < expected) outBuf[outPos++] = (byte)code;
                    prev = code;
                    continue;
                }

                int emit; bool kwk = false;
                if (code < nextCode) emit = code;
                else if (code == nextCode) { emit = prev; kwk = true; }
                else return null;

                int sp = 0, c = emit;
                while (c >= 0)
                {
                    if (sp >= stack.Length) return null;
                    stack[sp++] = suffix[c];
                    c = prefix[c];
                }
                byte first = stack[sp - 1];
                for (int i = sp - 1; i >= 0 && outPos < expected; i--) outBuf[outPos++] = stack[i];
                if (kwk && outPos < expected) outBuf[outPos++] = first;

                if (nextCode < MaxCodes)
                {
                    prefix[nextCode] = prev;
                    suffix[nextCode] = first;
                    nextCode++;
                    if (nextCode == (1 << codeWidth) - 1 && codeWidth < 12) codeWidth++;   // early change
                }
                prev = code;
                if (outPos >= expected) break;
            }
            return outBuf;
        }

        private static int ReadCodeMsb(byte[] s, ref int bitPos, int width, int endBit)
        {
            if (bitPos + width > endBit) return -1;
            int code = 0;
            for (int i = 0; i < width; i++)
            {
                int bytePos = bitPos >> 3;
                int bit = (s[bytePos] >> (7 - (bitPos & 7))) & 1;      // most-significant bit first
                code = (code << 1) | bit;
                bitPos++;
            }
            return code;
        }

        private static byte[] Slice(byte[] d, int off, int len)
        {
            var s = new byte[len];
            System.Array.Copy(d, off, s, 0, len);
            return s;
        }

        private static bool ReadHeader(byte[] d, out bool bigEndian, out int ifd0)
        {
            bigEndian = false; ifd0 = 0;
            if (d == null || d.Length < 8) return false;
            if (d[0] == 'I' && d[1] == 'I') bigEndian = false;
            else if (d[0] == 'M' && d[1] == 'M') bigEndian = true;
            else return false;
            if (ReadU16(d, bigEndian, 2) != 42) return false;
            long off = ReadU32(d, bigEndian, 4);
            if (off < 8 || off > d.Length) return false;
            ifd0 = (int)off;
            return true;
        }

        private static long ReadValue(byte[] d, bool be, int type, int fieldOffset)
        {
            var v = ReadValues(d, be, type, 1, fieldOffset);
            return v.Length > 0 ? v[0] : 0;
        }

        private static long[] ReadValues(byte[] d, bool be, int type, int count, int fieldOffset)
        {
            int ts = TypeSize(type);
            if (ts == 0 || count <= 0 || count > (1 << 24)) return new long[0];
            long total = (long)ts * count;
            int dataOffset = total <= 4 ? fieldOffset : (int)ReadU32(d, be, fieldOffset);
            var vals = new long[count];
            for (int i = 0; i < count; i++)
            {
                int p = dataOffset + i * ts;
                if (p < 0 || p + ts > d.Length) { vals[i] = 0; continue; }
                switch (ts)
                {
                    case 1: vals[i] = d[p]; break;
                    case 2: vals[i] = ReadU16(d, be, p); break;
                    case 4: vals[i] = ReadU32(d, be, p); break;
                    default: vals[i] = 0; break;
                }
            }
            return vals;
        }

        private static int TypeSize(int type)
        {
            switch (type)
            {
                case 1: case 2: case 6: case 7: return 1;   // BYTE / ASCII / SBYTE / UNDEFINED
                case 3: case 8: return 2;                   // SHORT / SSHORT
                case 4: case 9: case 11: return 4;          // LONG / SLONG / FLOAT
                default: return 0;                          // RATIONAL/DOUBLE etc. not needed here
            }
        }

        private static int ReadU16(byte[] d, bool be, int i)
            => be ? (d[i] << 8) | d[i + 1] : d[i] | (d[i + 1] << 8);

        private static long ReadU32(byte[] d, bool be, int i)
            => be ? ((long)d[i] << 24) | ((long)d[i + 1] << 16) | ((long)d[i + 2] << 8) | d[i + 3]
                  : (long)d[i] | ((long)d[i + 1] << 8) | ((long)d[i + 2] << 16) | ((long)d[i + 3] << 24);
    }
}