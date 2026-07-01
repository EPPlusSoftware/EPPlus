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
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics
{
    internal abstract class GenericFontMetricsTextMeasurerBase
    {
        private FontScaleFactors _fontScaleFactors = new FontScaleFactors();
        private static Dictionary<uint, SerializedFontMetrics> _fonts;
        private static object _syncRoot = new object();

        public GenericFontMetricsTextMeasurerBase()
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

        internal protected bool IsValidFont(uint fontKey)
        {
            return _fonts.ContainsKey(fontKey);
        }

        internal protected TextMeasurement MeasureTextInternal(string text, uint fontKey, MeasurementFontStyles style, float size, bool wrapText = false, float lastLinePadding = 0f)
        {
            var sFont = _fonts[fontKey];
            var width = 0f;
            var maxWidth = 0f;
            var widthEA = 0f;
            var maxWidthEA = 0f;
            for (var x = 0; x < text.Length; x++)
            {
                var fnt = sFont;
                var c = text[x];
                
                // When word-wrap is enabled, split lines on explicit newlines (\n, \r) 
                // as well as soft wrap boundaries (spaces, tabs, and hyphens) as Excel does.
                if(wrapText && (c=='\n' || c=='\r' || c==' ' || c=='-' || c=='\t'))
                {
                    if(x>0 && c=='\r' && text[x-1]=='\n')
                    {
                        continue; //CRLF should be handled as one new line.
                    }
                    
                    // Hyphens remain visible at the end of the wrapped line.
                    // Therefore, we measure and add the hyphen width to the current line width before wrapping.
                    if (c == '-')
                    {
                        if (sFont.CharMetrics.ContainsKey(c))
                        {
                            width += fnt.ClassWidths[sFont.CharMetrics[c]];
                        }
                    }
                    if ((width + widthEA) > (maxWidth + maxWidthEA))
                    {
                        maxWidth = width;
                        maxWidthEA = widthEA;
                    }
                    
                    // Reset current line width and East Asian width for the next line
                    width = 0;
                    widthEA = 0;
                    
                    // Since hyphen is already counted, continue directly to avoid processing it again
                    if (c == '-')
                    {
                        continue;
                    }
                }

                //If east Asian char use default regardless of actual font.
                if (IsEastAsianChar(c))
                {
                    widthEA += GetEastAsianCharWidth(c, style);
                }
                else
                {
                    if (sFont.CharMetrics.ContainsKey(c))
                    {
                        var fw = fnt.ClassWidths[sFont.CharMetrics[c]];
                        if (Char.IsDigit(c)) fw *= FontScaleFactors.DigitsScalingFactor;
                        width += fw;
                    }
                    else if (char.IsControl(c)==false)
                    {
                        width += sFont.ClassWidths[fnt.DefaultWidthClass];
                    }
                }
            }
            
            // Only perform separate line measurements if we have active AutoFilter padding AND the text has actually wrapped into multiple lines.
            if (lastLinePadding > 0f && (maxWidth > 0f || maxWidthEA > 0f))
            {
                // Measure the preceding lines and the last line separately.
                // Apply the padding ONLY to the last line.
                float maxPreceding = maxWidth * size;
                maxPreceding *= _fontScaleFactors.GetScaleFactor(fontKey, maxPreceding);
                maxPreceding += maxWidthEA * size;

                float lastLine = width * size;
                lastLine *= _fontScaleFactors.GetScaleFactor(fontKey, lastLine);
                lastLine += widthEA * size;

                width = Math.Max(maxPreceding, lastLine + lastLinePadding);
            }
            else
            {
                if ((maxWidth + maxWidthEA) > (width + widthEA))
                {
                    width = maxWidth;
                    widthEA = maxWidthEA;
                }
                width *= size;
                widthEA *= size;
                var sf = _fontScaleFactors.GetScaleFactor(fontKey, width);
                width *= sf;
                width += widthEA;
                width += lastLinePadding;
            }

            var height = sFont.LineHeight1em * size;
            return new TextMeasurement(width, height);
        }

        static Dictionary<char, uint> AlphabetChars = new Dictionary<char, uint>
        {
            {'a', 0x06 },
            {'b', 0x07 },
            {'c', 0x05 },
            {'d', 0x07 },
            {'e', 0x06 },
            {'f', 0x04 },
            {'g', 0x07 },
            {'h', 0x07 },
            {'i', 0x03 },
            {'j', 0x03 },
            {'k', 0x06 },
            {'l', 0x03 },
            {'m', 0x09 },
            {'n', 0x07 },
            {'o', 0x07 },
            {'p', 0x07 },
            {'q', 0x07 },
            {'r', 0x04 },
            {'s', 0x05 },
            {'t', 0x04 },
            {'u', 0x07 },
            {'v', 0x05 },
            {'w', 0x09 },
            {'x', 0x05 },
            {'y', 0x05 },
            {'z', 0x05 },
            {'A', 0x07 },
            {'B', 0x06 },
            {'C', 0x07 },
            {'D', 0x08 },
            {'E', 0x06 },
            {'F', 0x06 },
            {'G', 0x08 },
            {'H', 0x08 },
            {'I', 0x03 },
            {'J', 0x04 },
            {'K', 0x06 },
            {'L', 0x05 },
            {'M', 0x0A },
            {'N', 0x08 },
            {'O', 0x09 },
            {'P', 0x06 },
            {'Q', 0x08 },
            {'R', 0x07 },
            {'S', 0x06 },
            {'T', 0x06 },
            {'U', 0x08 },
            {'V', 0x07 },
            {'W', 0x0B },
            {'X', 0x06 },
            {'Y', 0x05 },
            {'Z', 0x06 }
        };

        internal List<uint> MeasureTextSpacingInternal(string text, uint fontKey, MeasurementFontStyles style, float size, float ppi = 108.73578912433f)
        {
            var sFont = _fonts[fontKey];
            var chars = text.ToCharArray();

            var spacingBuffer = new List<uint>();

            var widthDefault = sFont.ClassWidths[sFont.DefaultWidthClass];

            float resolutionDifference = ppi / 96f;
            float ptSize = size * (72f / 96f);

            float finalFactor = resolutionDifference * ptSize;

            for (var x = 0; x < chars.Length; x++)
            {
                var fnt = sFont;
                var c = chars[x];

                var fntClass = sFont.CharMetrics.ContainsKey(c) ? sFont.CharMetrics[c] : fnt.DefaultWidthClass;
                float adjustmentFactor = 0.012f * ptSize * ((int)fntClass);

                float deviceUnits = fnt.ClassWidths[fntClass] * finalFactor - adjustmentFactor;

                var rounded = Math.Round(deviceUnits * 10, MidpointRounding.AwayFromZero);
                var final = rounded / 10;

                uint simplifiedWidth = (uint)(Math.Round(deviceUnits, MidpointRounding.AwayFromZero));
                spacingBuffer.Add(simplifiedWidth);
            }

            return spacingBuffer;
        }

        internal uint MeasureCharacter(char c, uint fontKey, MeasurementFontStyles style, float ptSize, float finalFactor)
        {
            var sFont = _fonts[fontKey];
            var fnt = sFont;

            var fntClass = sFont.CharMetrics.ContainsKey(c) ? sFont.CharMetrics[c] : fnt.DefaultWidthClass;
            float adjustmentFactor = 0.012f * ptSize * ((int)fntClass);

            float deviceUnits = fnt.ClassWidths[fntClass] * finalFactor - adjustmentFactor;

            uint simplifiedWidth = (uint)Math.Round(deviceUnits, MidpointRounding.AwayFromZero);
            return simplifiedWidth;
        }

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

        private static float GetEastAsianCharWidth(int cc, MeasurementFontStyles style)
        {
            var emWidth = (cc >= 65377 && cc <= 65439) ? 0.5f : 1f;
            if ((style & MeasurementFontStyles.Bold) != 0)
            {
                emWidth *= 1.05f;
            }
            return emWidth * (96F / 72F) * FontScaleFactors.JapaneseKanjiDefaultScalingFactor;
        }

        private static bool IsEastAsianChar(char c)
        {
            var cc = (int)c;

            return UniCodeRange.JapaneseKanji.Any(x => x.IsInRange(cc));
        }

    }
}
