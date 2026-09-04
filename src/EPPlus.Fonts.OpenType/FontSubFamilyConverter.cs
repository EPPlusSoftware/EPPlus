/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2026         EPPlus Software AB           Extracted from OpenTypeFontEngine
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Maps between <see cref="FontSubFamily"/> and <see cref="MeasurementFontStyles"/>.
    /// Neither type has anything to do with the font engine, so the mapping does not live there.
    /// </summary>
    public static class FontSubFamilyConverter
    {
        public static FontSubFamily ToSubFamily(MeasurementFontStyles style)
        {
            bool bold = (style & MeasurementFontStyles.Bold) == MeasurementFontStyles.Bold;
            bool italic = (style & MeasurementFontStyles.Italic) == MeasurementFontStyles.Italic;

            if (bold && italic)
                return FontSubFamily.BoldItalic;
            if (bold)
                return FontSubFamily.Bold;
            if (italic)
                return FontSubFamily.Italic;

            return FontSubFamily.Regular;
        }

        public static MeasurementFontStyles ToStyles(FontSubFamily subFamily)
        {
            switch (subFamily)
            {
                case FontSubFamily.Bold:
                    return MeasurementFontStyles.Bold;
                case FontSubFamily.Italic:
                    return MeasurementFontStyles.Italic;
                case FontSubFamily.BoldItalic:
                    return MeasurementFontStyles.Bold | MeasurementFontStyles.Italic;
                default:
                    // MeasurementFontStyles.Regular if that member exists; otherwise the
                    // zero value, which is what Regular means for a flags enum.
                    return default(MeasurementFontStyles);
            }
        }
    }
}