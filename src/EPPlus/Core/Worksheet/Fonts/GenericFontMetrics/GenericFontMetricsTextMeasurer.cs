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

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal class GenericFontMetricsTextMeasurer : GenericFontMetricsTextMeasurerBase, IWrapTextMeasurer
    {
        /// <summary>
        /// If the text measurer should measure wrap text cells during AutoFit calculations.
        /// Line breaks are considered on explicit newlines (CR, LF, CRLF) as well as soft wrap boundaries (spaces, tabs, and hyphens).
        /// </summary>
        public bool MeasureWrappedTextCells
        {
            get;
            set;
        }

        /// <summary>
        /// Measures the supplied text
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="font">Font of the text to measure</param>
        /// <returns>A <see cref="TextMeasurement"/></returns>
        public TextMeasurement MeasureText(string text, MeasurementFont font)
        {
            return MeasureText(text, font, MeasureWrappedTextCells);
        }

        /// <summary>
        /// Measures the supplied text with cell-specific wrapping and last-line padding
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="font">Font of the text to measure</param>
        /// <param name="wrapText">Whether word wrap is enabled for this cell</param>
        /// <param name="lastLinePadding">Padding in pixels to add only to the last wrapped line (e.g. for autofilter arrows)</param>
        /// <returns>A <see cref="TextMeasurement"/></returns>
        public TextMeasurement MeasureText(string text, MeasurementFont font, bool wrapText, float lastLinePadding = 0f)
        {
            var fontKey = GetKey(font.FontFamily, font.Style);
            if (!IsValidFont(fontKey)) return TextMeasurement.Empty;
            return MeasureTextInternal(text, fontKey, font.Style, font.Size, wrapText, lastLinePadding);
        }

        public bool ValidForEnvironment()
        {
            return true;
        }

        internal List<uint> MeasureIndividualCharacters(string text, MeasurementFont font, float ppi = 108.73578912433f)
        {
            var fontKey = GetKey(font.FontFamily, font.Style);
            if (IsValidFont(fontKey))
            {
                return MeasureTextSpacingInternal(text, fontKey, font.Style, font.Size, ppi);
            }
            else
            {
                throw new InvalidOperationException("Font is not valid");
            }
        }
        internal uint MeasureIndividualCharacter(char c, MeasurementFont font, float ppi = 108.73578912433f)
        {
            var fontKey = GetKey(font.FontFamily, font.Style);
            if (IsValidFont(fontKey))
            {
                float resolutionDifference = ppi / 96f;
                float ptSize = font.Size * (72f / 96f);
                float finalFactor = resolutionDifference * ptSize;

                return MeasureCharacter(c, fontKey, font.Style, ptSize, finalFactor);
            }
            else
            {
                throw new InvalidOperationException("Font is not valid");
            }
        }
    }
}
