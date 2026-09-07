using OfficeOpenXml.Packaging.Ionic.Zlib;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.Helpers
{
    internal static class PdfFlate
    {
        internal static byte[] Compress(byte[] data)
        {
            using (var ms = new MemoryStream())
            {
                using (var zs = new ZlibStream(ms, CompressionMode.Compress, CompressionLevel.BestCompression))
                {
                    zs.Write(data, 0, data.Length);
                }
                return ms.ToArray();
            }
        }

        internal static byte[] CompressLeaveOpen(byte[] data)
        {
            using (var output = new MemoryStream())
            {
                using (var z = new ZlibStream(output, CompressionMode.Compress, CompressionLevel.BestCompression, true))
                {
                    z.Write(data, 0, data.Length);
                }
                return output.ToArray();
            }
        }

        internal static byte[] Decompress(byte[] data)
        {
            using (var input = new MemoryStream(data))
            using (var z = new ZlibStream(input, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                byte[] buffer = new byte[8192];
                int n;
                while ((n = z.Read(buffer, 0, buffer.Length)) > 0) output.Write(buffer, 0, n);
                return output.ToArray();
            }
        }
    }
}
