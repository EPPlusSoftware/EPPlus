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
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfImageXObject : PdfObject
    {
        private readonly byte[] _bytes;
        internal int Width { get; }
        internal int Height { get; }
        internal string ColorSpace { get; }
        internal int BitsPerComponent { get; }
        internal string Filter { get; }
        internal string Decode { get; private set; }
        internal string DecodeParms { get; }

        internal bool HasSoftMask { get; }
        internal byte[] SoftMaskData { get; }
        internal int SoftMaskObjectNumber { get; set; }

        public PdfImageXObject(int objectNumber, byte[] imageBytes, int version = 0)
            : base(objectNumber, version)
        {
            if (IsJpeg(imageBytes))
            {
                _bytes = imageBytes;
                // A JPEG embeds verbatim: /DCTDecode is exactly the JPEG's own coding.
                Filter = "DCTDecode";
                BitsPerComponent = 8;
                ReadJpegInfo(imageBytes, out int width, out int height, out int components, out bool adobe);
                Width = width;
                Height = height;
                if (components == 4)
                {
                    ColorSpace = "/DeviceCMYK";
                    // Adobe writes CMYK JPEGs with every channel inverted; flip them back so the
                    // picture doesn't render as a negative. Straight (non-Adobe) CMYK is left as-is.
                    if (adobe) Decode = "[ 1 0 1 0 1 0 1 0 ]";
                }
                else
                {
                    ColorSpace = components == 1 ? "/DeviceGray" : "/DeviceRGB";
                }
            }
            else if (IsPng(imageBytes))
            {
                ReadPngHeader(imageBytes, out int width, out int height, out int bitDepth, out int colorType, out int _);
                Width = width;
                Height = height;
                Filter = "FlateDecode";
                if (colorType == 6 || colorType == 4)
                {
                    // Alpha channel present: decode the PNG and split colour from alpha. The colour
                    // samples become this image; the alpha rides along as a grayscale soft mask.
                    BitsPerComponent = 8;
                    ColorSpace = colorType == 6 ? "/DeviceRGB" : "/DeviceGray";
                    DecodePngWithAlpha(imageBytes, width, height, colorType, out byte[] color, out byte[] alpha);
                    _bytes = color;            // raw colour samples, re-deflated (no PNG predictor)
                    SoftMaskData = alpha;      // raw alpha, re-deflated -> companion /SMask object
                    HasSoftMask = true;
                }
                else
                {
                    // Opaque (0/2/3): keep the compressed pixel data as-is. The concatenated IDAT is a
                    // complete zlib stream of PNG-filtered rows — exactly what /FlateDecode + a PNG
                    // predictor expect — so the viewer does the inflate and un-filter for us.
                    BitsPerComponent = bitDepth;
                    _bytes = ReadPngIdat(imageBytes, out byte[] palette);
                    int colors;
                    switch (colorType)
                    {
                        case 0:   // greyscale
                            ColorSpace = "/DeviceGray";
                            colors = 1;
                            break;
                        case 3:   // palette index -> RGB lookup table carried inline
                            int hival = palette == null || palette.Length < 3 ? 0 : (palette.Length / 3) - 1;
                            ColorSpace = "[ /Indexed /DeviceRGB " + hival + " <" + ToHex(palette) + "> ]";
                            colors = 1;
                            break;
                        default:  // colour type 2 (truecolour RGB)
                            ColorSpace = "/DeviceRGB";
                            colors = 3;
                            break;
                    }
                    // Predictor 15 = "PNG optimum" (any of the five row filters), described by the
                    // pixel layout so the viewer can reverse the per-row filtering.
                    DecodeParms = "<< /Predictor 15 /Colors " + colors +
                                  " /BitsPerComponent " + bitDepth +
                                  " /Columns " + width + " >>";
                }
            }
            else
            {
                // Unsupported encodings are screened out in PrecomputeImages; keep a safe default so
                // an unexpected byte stream can't crash the export (future formats add a branch above).
                _bytes = imageBytes;
                Filter = "DCTDecode";
                BitsPerComponent = 8;
                ColorSpace = "/DeviceRGB";
            }

        }

        internal static bool CanEmbed(byte[] imageBytes)
        {
            if (IsJpeg(imageBytes)) return true;
            if (IsPng(imageBytes))
            {
                if (!ReadPngHeader(imageBytes, out int _, out int _, out int _, out int colorType, out int interlace))
                    return false;
                if (interlace != 0) return false;
                return colorType == 0 || colorType == 2 || colorType == 3;
            }
            return false;
        }

        private static bool IsJpeg(byte[] d) => d != null && d.Length > 2 && d[0] == 0xFF && d[1] == 0xD8;

        private static readonly byte[] _pngSignature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        private static bool IsPng(byte[] d)
        {
            if (d == null || d.Length<_pngSignature.Length) return false;
            for (int i = 0; i<_pngSignature.Length; i++)
                if (d[i] != _pngSignature[i]) return false;
            return true;
        }

        private string DictHeader()
        {
            string decode = string.IsNullOrEmpty(Decode) ? "" : $" /Decode {Decode}";
            string decodeParms = string.IsNullOrEmpty(DecodeParms) ? "" : $" /DecodeParms {DecodeParms}";
            return "<< /Type /XObject /Subtype /Image" +
                   $" /Width {Width} /Height {Height}" +
                   $" /ColorSpace {ColorSpace} /BitsPerComponent {BitsPerComponent}" +
                   decode +
                   $" /Filter /{Filter}" + decodeParms +
                   $" /Length {_bytes.Length} >>";
        }

        internal override string RenderDictionary()
        {
            // Debug/text dump only — never the real output — so the binary body is elided.
            return DictHeader() + $"\nstream\n<{_bytes.Length} bytes of image data>\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, DictHeader() + "\nstream\n");
            bw.Write(_bytes);                 // raw JPEG — not Flate-compressed (already DCT-coded)
            WriteAscii(bw, "\nendstream");
        }

        // Minimal JPEG reader: walk the marker segments to the Start-Of-Frame and read the frame's
        // height, width and component count. Handles baseline and progressive SOFs.
        private static void ReadJpegInfo(byte[] d, out int width, out int height, out int components, out bool adobe)
        {
            width = 0; height = 0; components = 3; adobe = false;
            if (d == null || d.Length < 4 || d[0] != 0xFF || d[1] != 0xD8) return;   // not a JPEG
            int i = 2;
            while (i + 1 < d.Length)
            {
                if (d[i] != 0xFF) { i++; continue; }
                byte marker = d[i + 1];
                if (marker == 0xFF) { i++; continue; }                               // fill byte
                // Standalone markers without a length: SOI, EOI, RSTn, TEM.
                if (marker == 0xD8 || marker == 0xD9 || (marker >= 0xD0 && marker <= 0xD7) || marker == 0x01)
                {
                    i += 2; continue;
                }
                if (i + 3 >= d.Length) return;
                int segLen = (d[i + 2] << 8) | d[i + 3];
                // Adobe APP14 marker (FF EE) with an "Adobe" payload: Adobe-written, so 4-channel
                // data is stored inverted (the caller adds /Decode to correct it). APP14 precedes SOF.
                if (marker == 0xEE && i + 8 < d.Length &&
                    d[i + 4] == (byte)'A' && d[i + 5] == (byte)'d' && d[i + 6] == (byte)'o' &&
                    d[i + 7] == (byte)'b' && d[i + 8] == (byte)'e')
                {
                    adobe = true;
                }
                // SOF markers hold the frame size: C0..CF except C4 (DHT), C8 (JPG ext), CC (DAC).
                if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
                {
                    if (i + 9 >= d.Length) return;
                    height = (d[i + 5] << 8) | d[i + 6];
                    width = (d[i + 7] << 8) | d[i + 8];
                    components = d[i + 9];
                    return;
                }
                if (segLen < 2) return;                                               // malformed
                i += 2 + segLen;
            }
        }

        private static bool ReadPngHeader(byte[] d, out int width, out int height, out int bitDepth, out int colorType, out int interlace)
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
        private static byte[] ReadPngIdat(byte[] d, out byte[] palette)
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

        private static readonly char[] _hex = "0123456789ABCDEF".ToCharArray();
        private static string ToHex(byte[] bytes)
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
        private static int ReadBE32(byte[] d, int i) => (d[i] << 24) | (d[i + 1] << 16) | (d[i + 2] << 8) | d[i + 3];
        private static string Ascii(byte[] d, int i, int len) => Encoding.ASCII.GetString(d, i, len);
    }
}
