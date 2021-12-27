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
 *************************************************************************************************/
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts.Serialization;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts
{
    internal static class TextMeasurer
    {
        private static Dictionary<uint, SerializedFontMetrics> _fonts;
        private static object _syncRoot = new object();

        private static void Initialize()
        {
            lock(_syncRoot)
            {
                if (_fonts == null)
                {
                    _fonts = FontMetricsLoader.LoadFontMetrics();
                }
            }
        }

        private static float MeasureText(string text, float sizeInEm, SerializedFontMetrics sFont)
        {
            var width = 0f;
            var chars = text.ToCharArray();
            for(var x = 0; x < chars.Length; x++)
            {
                var c = chars[x];
                if(sFont.AdvanceWidths.ContainsKey(c))
                {
                    width += sFont.AdvanceWidths[c];
                }
                else
                {
                    width += sFont.DefaultAdvanceWidth;
                }
                if (x < chars.Length - 1 && sFont.KerningPairs != null)
                {
                    var nextChar = chars[x + 1];
                    var pairKey = $"{c}.{nextChar}";
                    if (sFont.KerningPairs.ContainsKey(pairKey))
                    {
                        width += sFont.KerningPairs[pairKey];
                    }
                }
            }
            var ems = width/sFont.UnitsPerEm;
            var emSize = ems * sizeInEm;
            var pixels = emSize * (96F/72F);
            var scaleFactor = FontScaleFactors.Instance[sFont.GetKey()];
            //return (float)pixels * (1f/scaleFactor) + (3.55555f * (8f/sizeInEm));
            return (float)pixels * (1f / scaleFactor);
        }

        public static float Measure(string text, float size, SerializedFontFamilies fontFamily, FontSubFamilies subFamily)
        {
            Initialize();
            var key = SerializedFontMetrics.GetKey(fontFamily, subFamily);
            if (!_fonts.ContainsKey(key)) return -1;
            return MeasureText(text, size, _fonts[key]);
        }

        public static float Measure(string text, float size, uint serializedFontMetricsId)
        {
            Initialize();
            if (!_fonts.ContainsKey(serializedFontMetricsId)) return -1;
            return MeasureText(text, size, _fonts[serializedFontMetricsId]);
        }
    }
}
