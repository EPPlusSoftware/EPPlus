using System;
using System.IO;
using OfficeOpenXml.PDF.PdfSettings;

namespace FontLab1.GenericMeasurements
{
    internal static class GenericFonts
    {
        internal static TtfFont GetFontData(PdfPageSettings pageSettings, string fontName, string subFamily = "Regular")
        {
            string DefaultFontPath = @"c:\Windows\Fonts";
            var basePath = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var userFontPath = Path.Combine(basePath, "Microsoft\\Windows\\Fonts");

            var factory = new GenericFontFactory(userFontPath);
            var fontData = factory.Create(fontName, subFamily);
            if (fontData == null)
            {
                return factory.Create(fontName, subFamily);
            }
            return fontData;
        }
    }
}
