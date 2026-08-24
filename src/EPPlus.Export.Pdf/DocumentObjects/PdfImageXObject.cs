using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfImageXObject : PdfObject
    {
        private readonly byte[] _bytes;
        internal int Width { get; }
        internal int Height { get; }
        internal string ColorSpace { get; }
        internal string Filter { get; }
        internal string Decode { get; private set; }

        public PdfImageXObject(int objectNumber, byte[] imageBytes, int version = 0)
            : base(objectNumber, version)
        {
            _bytes = imageBytes;
            if (IsJpeg(imageBytes))
            {
                // A JPEG embeds verbatim: /DCTDecode is exactly the JPEG's own coding.
                Filter = "DCTDecode";
                ReadJpegInfo(imageBytes, out int width, out int height, out int components, out bool adobe);
                Width = width;
                Height = height;
                if (components == 4)
                {
                    ColorSpace = "DeviceCMYK";
                    // Adobe writes CMYK JPEGs with every channel inverted; flip them back so the
                    // picture doesn't render as a negative. Straight (non-Adobe) CMYK is left as-is.
                    if (adobe) Decode = "[ 1 0 1 0 1 0 1 0 ]";
                }
                else
                {
                    ColorSpace = components == 1 ? "DeviceGray" : "DeviceRGB";
                }
            }
            else
            {
                // Unsupported encodings are screened out in PrecomputeImages; keep a safe default so
                // an unexpected byte stream can't crash the export (future formats add a branch above).
                Filter = "DCTDecode";
                ColorSpace = "DeviceRGB";
            }
        }

        private static bool IsJpeg(byte[] d) => d != null && d.Length > 2 && d[0] == 0xFF && d[1] == 0xD8;

        private string DictHeader()
        {
            return "<< /Type /XObject /Subtype /Image" +
                   $" /Width {Width} /Height {Height}" +
                   $" /ColorSpace /{ColorSpace} /BitsPerComponent 8" +
                   $" /Filter /{Filter} /Length {_bytes.Length} >>";
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
    }
}
