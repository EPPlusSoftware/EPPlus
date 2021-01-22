using Ionic.Zip;
using OfficeOpenXml.Packaging.Ionic.Zip;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.Core
{
    public static class ZipHelper
    {
        internal static string UncompressEntry(ZipInputStream zipStream, ZipEntry entry)
        {
            var content = new byte[entry.UncompressedSize];
            var size = zipStream.Read(content, 0, (int)entry.UncompressedSize);
            return Encoding.UTF8.GetString(content);
        }

        internal static ZipInputStream OpenZipResource()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var templateStream = assembly.GetManifestResourceStream("OfficeOpenXml.resources.DefaultTableStyles.cst");
            var zipStream = new ZipInputStream(templateStream);
            return zipStream;
        }
    }
}
