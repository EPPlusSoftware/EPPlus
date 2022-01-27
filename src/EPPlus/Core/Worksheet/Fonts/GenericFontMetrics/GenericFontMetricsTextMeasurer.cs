using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal class GenericFontMetricsTextMeasurer : ITextMeasurer
    {
        private static Dictionary<uint, SerializedFontMetrics> _fonts;
        private static object _syncRoot = new object();

        public GenericFontMetricsTextMeasurer()
        {
            Initialize();
        }

        private static void Initialize()
        {
            lock (_syncRoot)
            {
                if (_fonts == null)
                {
                    _fonts = GenericFontMetricsLoader.LoadFontMetrics();
                }
            }
        }

        public TextMeasurement MeasureText(string text, ExcelFont font)
        {
            var fontKey = GetKey(font.FontFamily, font.Style);
            if (!_fonts.ContainsKey(fontKey)) return TextMeasurement.Empty;
            var sFont = _fonts[fontKey];
            var width = 0f;
            var chars = text.ToCharArray();
            for (var x = 0; x < chars.Length; x++)
            {
                var fnt = sFont;
                var c = chars[x];
                // if east asian char use default ea font (MS Gothic) regardless of actual font.
                if (IsEastAsianChar(c))
                {
                    width += GetEastAsianCharWidth(c, font.Style);
                }
                else
                {
                    if (sFont.CharMetrics.ContainsKey(c))
                    {
                        width += fnt.ClassWidths[sFont.CharMetrics[c]];
                    }
                    else
                    {
                        width += fnt.DefaultWidth1em;
                    }
                }
                
            }
            width *= font.Size;
            var sf = FontScaleFactors.GetScaleFactor(fontKey, width);
            width *= sf;
            var height = sFont.LineHeight1em * font.Size;
            return new TextMeasurement(width, height);
        }

        public static uint GetKey(FontMetricsFamilies family, FontSubFamilies subFamily)
        {
            var k1 = (ushort)family;
            var k2 = (ushort)subFamily;
            return (uint)((k1 << 16) | ((k2) & 0xffff));
        }

        public static uint GetKey(string fontFamily, FontStyles fontStyle)
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

        private static float GetEastAsianCharWidth(int cc, FontStyles style)
        {
            var emWidth = (cc >= 65377 && cc <= 65439) ? 0.5f : 1f;
            if ((style & FontStyles.Bold) != 0)
            {
                emWidth *= 1.05f;
            }
            return emWidth * (96F / 72F);
        }

        private static bool IsEastAsianChar(char c)
        {
            var cc = (int)c;

            return UniCodeRange.JapaneseKanji.Any(x => x.IsInRange(cc));
        }
    }
}
