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
  09/01/2026         EPPlus Software AB       Resolve missing subfamilies to Regular
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.GenericFontWidths;
using OfficeOpenXml.Core.Worksheet.Fonts.GenericFontMetrics;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements
{
    internal class GenericFontMetricsTextMeasurer : GenericFontMetricsTextMeasurerBase, ITextMeasurer
    {
        /// <summary>
        /// If the text measurer should measure wrap text cells. 
        /// Only CR, LF or CRLF should be considered.
        /// </summary>
#pragma warning disable 618
        public bool MeasureWrappedTextCells
        {
            get;
            set;
        }
#pragma warning restore 618
        /// 
        /// <summary>
        /// 
        /// </summary>
        public eWrappedTextAutofitMode WrappedTextAutofitMode
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
            // ResolveKey rather than GetKey: fifteen family/subfamily combinations have no
            // metrics because the font is not shipped by Windows at all, and those now fall
            // back to the family's Regular instead of measuring as zero width.
            var fontKey = GenericTextMeasurerKey.ResolveKey(font.FontFamily, font.Style);
            if (fontKey == uint.MaxValue) return TextMeasurement.Empty;

            // The original style is still passed through: it drives the East Asian width
            // handling, which does not depend on which file the metrics came from.
            return MeasureTextInternal(text, fontKey, font.Style, font.Size, WrappedTextAutofitMode);
        }

        public bool ValidForEnvironment()
        {
            return true;
        }

        internal List<uint> MeasureIndividualCharacters(string text, MeasurementFont font, float ppi = 108.73578912433f)
        {
            var fontKey = GenericTextMeasurerKey.ResolveKey(font.FontFamily, font.Style);
            if (fontKey == uint.MaxValue)
            {
                throw new InvalidOperationException(
                    string.Format("No font metrics available for {0} {1}", font.FontFamily, font.Style));
            }
            return MeasureTextSpacingInternal(text, fontKey, font.Style, font.Size, ppi);
        }

        internal uint MeasureIndividualCharacter(char c, MeasurementFont font, float ppi = 108.73578912433f)
        {
            var fontKey = GenericTextMeasurerKey.ResolveKey(font.FontFamily, font.Style);
            if (fontKey == uint.MaxValue)
            {
                throw new InvalidOperationException(
                    string.Format("No font metrics available for {0} {1}", font.FontFamily, font.Style));
            }

            float resolutionDifference = ppi / 96f;
            float ptSize = font.Size * (72f / 96f);
            float finalFactor = resolutionDifference * ptSize;

            return MeasureCharacter(c, fontKey, font.Style, ptSize, finalFactor);
        }
    }
}