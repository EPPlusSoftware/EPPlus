using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Integration
{
    public class TextHandler
    {
        internal float CurrentFontSize { get; private set; }

        TextShaper _currentShaper;
        TextLayoutEngine _currentLayout;

        public TextHandler(MeasurementFont mf) 
        {
            CurrentFontSize = mf.Size;
            SetFont(mf);
        }

        public void SetFontSize(float fontSize)
        {
            CurrentFontSize = fontSize;
        }

        public void SetFont(MeasurementFont mf)
        {
            CurrentFontSize = mf.Size;
            _currentShaper = (TextShaper)OpenTypeFonts.GetShaperForFont(mf);
            _currentLayout = OpenTypeFonts.GetTextLayoutEngineForFont(mf);
        }

        /// <summary>
        /// Gets single line spacing (baseline-to-baseline distance).
        /// </summary>
        public float GetLineHeightInPoints()
        {
            return _currentShaper.GetLineHeightInPoints(CurrentFontSize);
        }

        /// <summary>
        /// Calculates the total height of the font, in points.
        /// </summary>
        public float GetFontHeightInPoints()
        {
            return _currentShaper.GetFontHeightInPoints(CurrentFontSize);
        }

        /// <summary>
        /// Calculates the distance from the top of the font's bounding box to the baseline.
        /// </summary>
        /// <param name="fontSize">The font size, in points, for which to calculate the baseline position. Must be a positive value.</param>
        /// <returns>The distance, in points, from the top of the font's bounding box to the baseline for the given font size.</returns>
        public float GetAscentInPoints()
        {
            return _currentShaper.GetAscentInPoints(CurrentFontSize);
        }

        /// <summary>
        /// Calculates the font descent in points.
        /// </summary>
        public float GetDescentInPoints()
        {
            return _currentShaper.GetDescentInPoints(CurrentFontSize);
        }

        /// <summary>
        /// Measures the width of text in PDF points.
        /// </summary>
        public float MeasureTextInPoints(string text, ShapingOptions options = null)
        {
            return _currentShaper.MeasureTextInPoints(text, CurrentFontSize, options);
        }

        /// <summary>
        /// Measures the width of text in pixels.
        /// </summary>
        public float MeasureTextInPixels(string text, float dpi=96, ShapingOptions options = null)
        {
            return _currentShaper.MeasureTextInPixels(text, CurrentFontSize, dpi, options);
        }

        /// <summary>
        /// Measure multi-line text and return bounding box.
        /// </summary>
        public MultiLineMetrics MeasureLines(string text, ShapingOptions options)
        {
            return _currentShaper.MeasureLines(text, CurrentFontSize, options);
        }

        public List<string> WrapText(
           string text,
           double maxWidthPoints,
           ShapingOptions options = null)
        {
            return _currentLayout.WrapText(text, CurrentFontSize, maxWidthPoints, 0, options);
        }
    }
}
