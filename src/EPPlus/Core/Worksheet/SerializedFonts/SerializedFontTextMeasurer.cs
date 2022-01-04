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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts.Serialization;
using OfficeOpenXml.Interfaces.Text;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.SerializedFonts
{
    internal class SerializedFontTextMeasurer : ITextMeasurer
    {
        private static Dictionary<uint, SerializedFontMetrics> _fonts;
        private static object _syncRoot = new object();

        public SerializedFontTextMeasurer()
        {
            Initialize();
        }

        private static void Initialize()
        {
            lock (_syncRoot)
            {
                if (_fonts == null)
                {
                    _fonts = FontMetricsLoader.LoadFontMetrics();
                }
            }
        }

        private static float FduToPixels(float sizeInEm, float fdu, ushort unitsPerEm)
        {
            var ems = fdu / unitsPerEm;
            var emSize = ems * sizeInEm;
            var pixels = emSize * (96F / 72F);
            return pixels;
        }

        public static uint GetKey(SerializedFontFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }

        public static uint GetKey(string fontFamily, FontStyles fontStyle)
        {
            var enumName = fontFamily.Replace(" ", string.Empty);
            var values = Enum.GetValues(typeof(SerializedFontFamilies));
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
            var family = (SerializedFontFamilies)Enum.Parse(typeof(SerializedFontFamilies), enumName);
            var subFamily = FontSubFamilies.Regular;
            switch (fontStyle)
            {
                case FontStyles.Bold:
                    subFamily = FontSubFamilies.Bold;
                    break;
                case FontStyles.Italic:
                    subFamily = FontSubFamilies.Italic;
                    break;
                case FontStyles.Italic | FontStyles.Bold:
                    subFamily = FontSubFamilies.BoldItalic;
                    break;
                default:
                    break;
            }
            return GetKey(family, subFamily);
        }

        //TextMeasurement ITextMeasurer.MeasureText(string text, string fontFamily, FontStyles fontStyle, float fontSize)
        TextMeasurement ITextMeasurer.MeasureText(string text, ExcelFont font)
        {
            var fontKey = GetKey(font.FontFamily, font.Style);
            if (!_fonts.ContainsKey(fontKey)) return TextMeasurement.Empty;
            var sFont = _fonts[fontKey];
            var width = 0f;
            var chars = text.ToCharArray();
            for (var x = 0; x < chars.Length; x++)
            {
                var c = chars[x];
                if (sFont.AdvanceWidths.ContainsKey(c))
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
            //var ems = width/sFont.UnitsPerEm;
            //var emSize = ems * sizeInEm;
            //var pixels = emSize * (96F/72F);
            //var scaleFactor = FontScaleFactors.Instance[sFont.GetKey()];
            //return (float)pixels * (1f/scaleFactor) + (3.55555f * (8f/sizeInEm));
            var pixelWidth = FduToPixels(font.Size, width, sFont.UnitsPerEm);
            var pixelHeight = FduToPixels(font.Size, sFont.LineHeight, sFont.UnitsPerEm);
            return new TextMeasurement(pixelWidth, pixelHeight);// * (1f / scaleFactor);
        }
    }
}
