using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText.Interfaces;

namespace OfficeOpenXml.Interfaces.RichText
{
    /// <summary>
    /// Data holder for basic font data
    /// </summary>
    public class FontDataBasic : IFontData
    {
        public FontDataBasic(string fontFamily, FontSubFamily subFamily, double fontSize) 
        {
            FontFamily = fontFamily;
            SubFamily = subFamily;
            FontSize = fontSize;
        }
        public FontDataBasic() {}
        public FontDataBasic(MeasurementFont font)
        {
            FontFamily = font.FontFamily;
            SubFamily = GetFontSubFamily(font.Style);
            FontSize = font.Size;
        }

        public string FontFamily { get; set; }
        public FontSubFamily SubFamily { get; set; }
        public double FontSize { get; set; }

        public static FontSubFamily GetFontSubFamily(MeasurementFontStyles style)
        {
            if ((style & (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic)) ==
                (MeasurementFontStyles.Bold | MeasurementFontStyles.Italic))
            {
                return FontSubFamily.BoldItalic;
            }
            else if ((style & MeasurementFontStyles.Bold) == MeasurementFontStyles.Bold)
            {
                return FontSubFamily.Bold;
            }
            else if ((style & MeasurementFontStyles.Italic) == MeasurementFontStyles.Italic)
            {
                return FontSubFamily.Italic;
            }

            return FontSubFamily.Regular;
        }
    }
}
