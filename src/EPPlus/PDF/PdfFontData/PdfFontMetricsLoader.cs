using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Packaging.Ionic.Zip;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.PDF.PdfFontData
{
    internal class PdfFontMetricsLoader
    {
        /// <summary>
        /// Loads all serialized font metrics from the resources/SerializedFonts.zip archive
        /// </summary>
        internal static Dictionary<uint, PdfFontProperties> LoadFontMetrics()
        {
            var fonts = new Dictionary<uint, PdfFontProperties>();
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream("OfficeOpenXml.resources.PdfTextMetrics.zip"))
            {
                var zipStream = new ZipInputStream(stream);
                ZipEntry entry;
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (!entry.IsDirectory && Path.GetExtension(entry.FileName) == ".fmtr")
                    {
                        var bytes = new byte[entry.UncompressedSize];
                        var size = zipStream.Read(bytes, 0, (int)entry.UncompressedSize);
                        using (var ms = RecyclableMemory.GetStream(bytes))
                        {
                            var fnt = PdfFontDataReader.Deserialize(ms);
                            fonts.Add(fnt.GetKey(), fnt);
                        }

                    }
                }
            }
            return fonts;
        }
    }
}
