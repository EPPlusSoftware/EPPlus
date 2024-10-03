﻿/*************************************************************************************************
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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal class GenericFontMetricsTextMeasurer : GenericFontMetricsTextMeasurerBase, ITextMeasurer
    {
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

        internal List<uint> MeasureIndividualCharacters(string text, MeasurementFont font)
        {
            var fontKey = GetKey(font.FontFamily, font.Style);
            if (IsValidFont(fontKey))
            {
                return MeasureTextSpacingInternal(text, fontKey, font.Style, font.Size);
            }
            else
            {
                throw new InvalidOperationException("Font is not valid");
            }
        }
    }
}
