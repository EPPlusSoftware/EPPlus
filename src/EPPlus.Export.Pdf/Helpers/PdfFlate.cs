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
using EPPlus.Export.Pdf.Settings;
using System.IO;

namespace EPPlus.Export.Pdf.Helpers
{
    internal static class PdfFlate
    {
        private static CompressionLevel GetCompressionLevel(PdfCompressionLevel level)
        {
            switch (level)
            {
                case PdfCompressionLevel.Fastest: return CompressionLevel.BestSpeed;
                case PdfCompressionLevel.Balanced: return CompressionLevel.Default;
                default: return CompressionLevel.BestCompression;
            }
        }

        internal static byte[] Compress(byte[] data, PdfCompressionLevel level = PdfCompressionLevel.Smallest)
        {
            using (var ms = new MemoryStream())
            {
                using (var zs = new ZlibStream(ms, CompressionMode.Compress, GetCompressionLevel(level)))
                {
                    zs.Write(data, 0, data.Length);
                }
                return ms.ToArray();
            }
        }

        internal static byte[] CompressLeaveOpen(byte[] data, PdfCompressionLevel level = PdfCompressionLevel.Smallest)
        {
            using (var output = new MemoryStream())
            {
                using (var z = new ZlibStream(output, CompressionMode.Compress, GetCompressionLevel(level), true))
                {
                    z.Write(data, 0, data.Length);
                }
                return output.ToArray();
            }
        }

        internal static byte[] Decompress(byte[] data, long maxBytes = long.MaxValue)
        {
            using (var input = new MemoryStream(data))
            using (var z = new ZlibStream(input, CompressionMode.Decompress))
            using (var output = new MemoryStream())
            {
                byte[] buffer = new byte[8192];
                int n;
                long total = 0;
                while ((n = z.Read(buffer, 0, buffer.Length)) > 0)
                {
                    total += n;
                    if (total > maxBytes)
                    {
                        throw new InvalidDataException("Compressed image stream expands beyond the allowed size.");
                    }
                    output.Write(buffer, 0, n);
                }
                return output.ToArray();
            }
        }
    }
}
