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
using System.Text;

namespace EPPlus.Export.Pdf.Helpers
{
    internal static class PngDecoder
    {
        private static readonly byte[] _pngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        internal static bool IsPng(byte[] d)
        {
            if (d == null || d.Length < _pngSignature.Length) return false;
            for (int i = 0; i < _pngSignature.Length; i++)
                if (d[i] != _pngSignature[i]) return false;
            return true;
        }

        internal static bool ReadPngHeader(byte[] d, out int width, out int height, out int bitDepth, out int colorType, out int interlace)
        {
            width = height = bitDepth = colorType = interlace = 0;
            if (!IsPng(d)) return false;
            int p = _pngSignature.Length;                        // first chunk starts after the signature
            if (p + 8 + 13 > d.Length) return false;
            if (Ascii(d, p + 4, 4) != "IHDR") return false;
            int q = p + 8;                                       // IHDR chunk data
            width = ReadBE32(d, q);
            height = ReadBE32(d, q + 4);
            bitDepth = d[q + 8];
            colorType = d[q + 9];
            // q+10 compression, q+11 filter (both always 0), q+12 interlace (0 none, 1 Adam7).
            interlace = d[q + 12];
            return true;
        }

        // Walk the chunk list and return the concatenated IDAT data (the zlib pixel stream) plus the
        // palette, if any. The zlib stream can be split across several IDAT chunks, so it is stitched
        // back together in order.
        internal static byte[] ReadPngIdat(byte[] d, out byte[] palette)
        {
            palette = null;
            using (var idat = new MemoryStream())
            {
                int p = _pngSignature.Length;
                while (p + 8 <= d.Length)
                {
                    int len = ReadBE32(d, p);
                    string type = Ascii(d, p + 4, 4);
                    int dataStart = p + 8;
                    if (len < 0 || dataStart + len + 4 > d.Length) break;   // truncated / malformed
                    if (type == "PLTE")
                    {
                        palette = new byte[len];
                        System.Array.Copy(d, dataStart, palette, 0, len);
                    }
                    else if (type == "IDAT")
                    {
                        idat.Write(d, dataStart, len);
                    }
                    else if (type == "IEND")
                    {
                        break;
                    }
                    p = dataStart + len + 4;                                 // skip data + 4-byte CRC
                }
                return idat.ToArray();
            }
        }

        internal static void DecodePngWithAlpha(byte[] pngBytes, int width, int height, int colorType, out byte[] deflatedColor, out byte[] deflatedAlpha)
        {
            int channels = colorType == 6 ? 4 : 2;             // RGBA or grey+alpha
            int colorChannels = colorType == 6 ? 3 : 1;
            byte[] filtered = PdfFlate.Decompress(ReadPngIdat(pngBytes, out byte[] _));

            int stride = width * channels;                     // 8-bit: one byte per channel
            var color = new byte[width * height * colorChannels];
            var alpha = new byte[width * height];
            var prev = new byte[stride];
            var cur = new byte[stride];
            int pos = 0, ci = 0, ai = 0;
            for (int y = 0; y < height; y++)
            {
                int filter = pos < filtered.Length ? filtered[pos++] : 0;   // per-row filter type byte
                for (int x = 0; x < stride; x++)
                {
                    int raw = pos < filtered.Length ? filtered[pos++] : 0;
                    int a = x >= channels ? cur[x - channels] : 0;   // reconstructed byte to the left
                    int b = prev[x];                                 // byte above
                    int c = x >= channels ? prev[x - channels] : 0;  // byte above-left
                    int val;
                    switch (filter)
                    {
                        case 1: val = raw + a; break;                        // Sub
                        case 2: val = raw + b; break;                        // Up
                        case 3: val = raw + ((a + b) >> 1); break;           // Average
                        case 4: val = raw + Paeth(a, b, c); break;           // Paeth
                        default: val = raw; break;                           // None
                    }
                    cur[x] = (byte)(val & 0xFF);
                }
                // De-interleave this row: colour bytes to the image, the last channel to the mask.
                for (int x = 0; x < width; x++)
                {
                    int p = x * channels;
                    if (colorType == 6)
                    {
                        color[ci++] = cur[p];
                        color[ci++] = cur[p + 1];
                        color[ci++] = cur[p + 2];
                        alpha[ai++] = cur[p + 3];
                    }
                    else
                    {
                        color[ci++] = cur[p];
                        alpha[ai++] = cur[p + 1];
                    }
                }
                var swap = prev; prev = cur; cur = swap;    // this row becomes "previous" for the next
            }
            deflatedColor = PdfFlate.CompressLeaveOpen(color);
            deflatedAlpha = PdfFlate.CompressLeaveOpen(alpha);
        }

        // PNG Paeth predictor (integer, no Math dependency).
        internal static int Paeth(int a, int b, int c)
        {
            int p = a + b - c;
            int pa = p > a ? p - a : a - p;
            int pb = p > b ? p - b : b - p;
            int pc = p > c ? p - c : c - p;
            if (pa <= pb && pa <= pc) return a;
            return pb <= pc ? b : c;
        }

        // zlib (RFC 1950) round-trips: PNG IDAT and PDF /FlateDecode are both zlib streams, so the
        // same codec decompresses the IDAT and compresses the split colour / alpha back.


        private static readonly char[] _hex = "0123456789ABCDEF".ToCharArray();
        internal static string ToHex(byte[] bytes)
        {
            if (bytes == null) return "";
            var sb = new StringBuilder(bytes.Length * 2);
            foreach (var b in bytes)
            {
                sb.Append(_hex[b >> 4]);
                sb.Append(_hex[b & 0x0F]);
            }
            return sb.ToString();
        }

        // Big-endian 32-bit read (PNG stores all integers most-significant byte first).
        internal static int ReadBE32(byte[] d, int i) => (d[i] << 24) | (d[i + 1] << 16) | (d[i + 2] << 8) | d[i + 3];
        internal static string Ascii(byte[] d, int i, int len) => Encoding.ASCII.GetString(d, i, len);
    }
}
