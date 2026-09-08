/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
  09/01/2026         EPPlus Software AB       Added ResolveKey with subfamily fallback
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;

namespace EPPlus.Fonts.OpenType.GenericFontWidths
{
    public static class GenericTextMeasurerKey
    {
        internal static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }

        internal static uint GetKey(string fontFamily, MeasurementFontStyles fontStyle)
        {
            var enumName = fontFamily.Replace(" ", string.Empty);
            var values = Enum.GetValues(typeof(FontMetricsFamilies));
            var supported = false;
            foreach (var enumVal in values)
            {
                if (enumVal.ToString() == enumName)
                {
                    supported = true;
                    break;
                }
            }
            if (!supported) return uint.MaxValue;
            var family = (FontMetricsFamilies)Enum.Parse(typeof(FontMetricsFamilies), enumName);
            var subFamily = FontSubFamilies.Regular;
            switch (fontStyle)
            {
                case MeasurementFontStyles.Bold:
                    subFamily = FontSubFamilies.Bold;
                    break;
                case MeasurementFontStyles.Italic:
                    subFamily = FontSubFamilies.Italic;
                    break;
                case MeasurementFontStyles.Italic | MeasurementFontStyles.Bold:
                    subFamily = FontSubFamilies.BoldItalic;
                    break;
                default:
                    break;
            }
            return GetKey(family, subFamily);
        }

        /// <summary>
        /// Like <see cref="GetKey(string, MeasurementFontStyles)"/>, but returns a key that
        /// actually has metrics behind it: if the requested subfamily is missing, the Regular
        /// subfamily of the same family is used. Returns uint.MaxValue when neither exists.
        ///
        /// Callers that go on to measure should use this rather than GetKey. See
        /// <see cref="GenericFontMetricsCache.ResolveFontKey"/> for why the substitution is
        /// needed and why it does not adjust anything for Bold.
        /// </summary>
        internal static uint ResolveKey(string fontFamily, MeasurementFontStyles fontStyle)
        {
            return GenericFontMetricsCache.ResolveFontKey(GetKey(fontFamily, fontStyle));
        }

        /// <summary>
        /// Overload for callers that already have the enum values.
        /// </summary>
        internal static uint ResolveKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            return GenericFontMetricsCache.ResolveFontKey(GetKey(family, subFamily));
        }
    }
}