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
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal class GenericFontMetricsTextMeasurer : GenericFontMetricsTextMeasurerBase, ITextMeasurer
    {
        public float GetScalingFactorRowHeight(MeasurementFont font)
        {
            if (string.IsNullOrEmpty(font.FontFamily))
            {
                return 1f;
            }
            switch (font.FontFamily)
            {
                case "Arial":
                    return 1.02f;
                case "Times New Roman":
                    return 1.15f;
                case "Liberation Serif":
                    return 1.2f;
                case "Verdana":
                    return 1f;
                case "Century Gothic":
                    return 0.95f;
                case "Courier New":
                    return 1.1f;
                case "Arial Black":
                    return 1.02f;
                case "Corbel":
                    return 1.05f;
                case "Trebuchet MS":
                    return 1.05f;
                case "Rockwell":
                    return 0.97f;
                case "Tw Cen MT":
                    return 1.12f;
                case "Tw Cen MT Condensed":
                    return 1.13f;
                default:
                    return 1f;
            }
        }

        /// <summary>
        /// Measures the supplied text
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="font">Font of the text to measure</param>
        /// <returns>A <see cref="TextMeasurement"/></returns>
        public TextMeasurement MeasureText(string text, MeasurementFont font)
        {
            var fontKey = GetKey(font.FontFamily, font.Style);
            if (!IsValidFont(fontKey)) return TextMeasurement.Empty;
            return MeasureTextInternal(text, fontKey, font.Style, font.Size);
        }

        public bool ValidForEnvironment()
        {
            return true;
        }
    }
}
