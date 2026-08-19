using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfImageXObject : PdfObject
    {
        private readonly byte[] _jpeg;
        internal int Width { get; }
        internal int Height { get; }
        internal string ColorSpace { get; }

        public PdfImageXObject(int objectNumber, byte[] jpegBytes, int version = 0)
            : base(objectNumber, version)
        {
            _jpeg = jpegBytes;
            ReadJpegInfo(jpegBytes, out int width, out int height, out int components);
            Width = width;
            Height = height;
            ColorSpace = components == 1 ? "DeviceGray" : components == 4 ? "DeviceCMYK" : "DeviceRGB";
        }

        private string DictHeader()
        {
            return "<< /Type /XObject /Subtype /Image" +
                   $" /Width {Width} /Height {Height}" +
                   $" /ColorSpace /{ColorSpace} /BitsPerComponent 8" +
                   $" /Filter /DCTDecode /Length {_jpeg.Length} >>";
        }

        internal override string RenderDictionary()
        {
            // Debug/text dump only — never the real output — so the binary body is elided.
            return DictHeader() + $"\nstream\n<{_jpeg.Length} bytes of JPEG data>\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, DictHeader() + "\nstream\n");
            bw.Write(_jpeg);                 // raw JPEG — not Flate-compressed (already DCT-coded)
            WriteAscii(bw, "\nendstream");
        }

        // Minimal JPEG reader: walk the marker segments to the Start-Of-Frame and read the frame's
        // height, width and component count. Handles baseline and progressive SOFs.
        private static void ReadJpegInfo(byte[] d, out int width, out int height, out int components)
        {
            width = 0; height = 0; components = 3;
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
                // SOF markers hold the frame size: C0..CF except C4 (DHT), C8 (JPG ext), CC (DAC).
                if (marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC)
                {
                    if (i + 9 >= d.Length) return;
                    height = (d[i + 5] << 8) | d[i + 6];
                    width = (d[i + 7] << 8) | d[i + 8];
                    components = d[i + 9];
                    return;
                }
                if (segLen < 2) return;
                i += 2 + segLen;
            }
        }
    }
}
