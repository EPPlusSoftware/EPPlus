/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  1/4/2021         EPPlus Software AB           EPPlus Interfaces 1.0
 *************************************************************************************************/

namespace OfficeOpenXml.Interfaces.Drawing.Text
{
    /// <summary>
    /// Interface for measuring width and height of texts.
    /// </summary>
    public interface ITextMeasurer
    {
        /// <summary>
        /// Should return true if the text measurer is valid for this environment. 
        /// </summary>
        /// <returns>True if the measurer can be used else false.</returns>
        bool ValidForEnvironment();
        /// <summary>
        /// Measures width and height of the parameter <paramref name="text"/>.
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="font">The <see cref="MeasurementFont">font</see> to measure</param>
        /// <returns></returns>
        TextMeasurement MeasureText(string text, MeasurementFont font);
        /// <summary>
        /// If the text measurer should measure wrap text cells. 
        /// Line breaks are considered on explicit newlines (CR, LF, CRLF) as well as soft wrap boundaries (spaces, tabs, and hyphens).
        /// </summary>
        bool MeasureWrappedTextCells { get; set; }
    }

    /// <summary>
    /// Extension interface for advanced text measurement including cell-specific wrapping and last line padding.
    /// </summary>
    public interface IWrapTextMeasurer : ITextMeasurer
    {
        /// <summary>
        /// Measures the supplied text with cell-specific wrapping and last-line padding
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="font">Font of the text to measure</param>
        /// <param name="wrapText">Whether word wrap is enabled for this cell</param>
        /// <param name="lastLinePadding">Padding in pixels to add only to the last wrapped line (e.g. for autofilter arrows)</param>
        /// <returns>A <see cref="TextMeasurement"/></returns>
        TextMeasurement MeasureText(string text, MeasurementFont font, bool wrapText, float lastLinePadding);
    }

    /// <summary>
    /// High-compatibility extension methods for <see cref="ITextMeasurer"/>
    /// </summary>
    public static class TextMeasurerExtensions
    {
        /// <summary>
        /// Measures the supplied text with cell-specific wrapping and last-line padding, using the advanced interface if supported, or falling back to interface-level properties.
        /// </summary>
        public static TextMeasurement MeasureText(this ITextMeasurer measurer, string text, MeasurementFont font, bool wrapText, float lastLinePadding)
        {
            var wrapMeasurer = measurer as IWrapTextMeasurer;
            if (wrapMeasurer != null)
            {
                return wrapMeasurer.MeasureText(text, font, wrapText, lastLinePadding);
            }
            
            var prev = measurer.MeasureWrappedTextCells;
            try
            {
                measurer.MeasureWrappedTextCells = wrapText;
                var measurement = measurer.MeasureText(text, font);
                measurement.Width += lastLinePadding;
                return measurement;
            }
            finally
            {
                measurer.MeasureWrappedTextCells = prev;
            }
        }
    }
}
