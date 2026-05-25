using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Interfaces.RichText;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Integration.DataHolders
{
    public class OpenTypeFontInfoBase : IFontFormatBase
    {
        public virtual string Family { get; set; } = "Archivo Narrow";
        public virtual FontSubFamily SubFamily { get; set; } = FontSubFamily.Regular;
        public virtual float Size { get; set; } = 11f;

        public OpenTypeFontInfoBase(string fontFamily, FontSubFamily subFamily, float fontSize)
        {
            Family = fontFamily;
            SubFamily = subFamily;
            Size = fontSize;
        }
        public OpenTypeFontInfoBase() { }

        /// <summary>
        /// Legacy constructor. Prefer to avoid with new implementations. To be removed after refactor
        /// </summary>
        /// <param name="font"></param>
        public OpenTypeFontInfoBase(MeasurementFont font)
        {
            Family = font.FontFamily;
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

        public void SetFont(IFontFormatBase font)
        {
            Family = font.Family;
            Size = font.Size;
            SubFamily = font.SubFamily;
        }
    }
}
