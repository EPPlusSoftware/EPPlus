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
    }
}
