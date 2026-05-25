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
    public class FontDataBasic : FontDataDefaults
    {
        public FontDataBasic(string fontFamily, FontSubFamily subFamily, float fontSize) 
        {
            FamilyName = fontFamily;
            SubFamily = subFamily;
            Size = fontSize;
        }
        public FontDataBasic() {}

        /// <summary>
        /// Legacy constructor. Prefer to avoid with new implementations. To be removed after refactor
        /// </summary>
        /// <param name="font"></param>
        public FontDataBasic(MeasurementFont font)
        {
            FamilyName = font.FontFamily;
            SubFamily = GetFontSubFamily(font.Style);
            Size = font.Size;
        }

        /// <summary>
        /// Utility for Legacy constructor. Prefer to avoid with new implementations. To be removed after refactor
        /// </summary>
        /// <param name="style"></param>
        /// <returns></returns>
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
