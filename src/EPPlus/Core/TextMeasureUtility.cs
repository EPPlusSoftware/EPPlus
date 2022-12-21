using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class TextMeasureUtility
    {
        public TextMeasureUtility()
        {
            if (FontSize.FontWidths.ContainsKey(FontSize.NonExistingFont))
            {
                FontSize.LoadAllFontsFromResource();
                _fontWidthDefault = FontSize.FontWidths[FontSize.NonExistingFont];
            }
            ResetFontCache();
        }

        Dictionary<float, short> _fontWidthDefault = null;
        Dictionary<int, MeasurementFont> _fontCache;
        ITextMeasurer _genericMeasurer = new GenericFontMetricsTextMeasurer();
        MeasurementFont _nonExistingFont = new MeasurementFont() { FontFamily = FontSize.NonExistingFont };

        internal void ResetFontCache()
        {
            _fontCache = new Dictionary<int, MeasurementFont>();
        }

        internal Dictionary<int, MeasurementFont> FontCache
        {
            get { return _fontCache; }
        }

        internal TextMeasurement MeasureString(string t, int fntID, ExcelTextSettings ts)
        {
            var measureCache = new Dictionary<ulong, TextMeasurement>();
            ulong key = ((ulong)((uint)t.GetHashCode()) << 32) | (uint)fntID;
            if (!measureCache.TryGetValue(key, out var measurement))
            {
                var measurer = ts.PrimaryTextMeasurer;
                var font = _fontCache[fntID];
                measurement = measurer.MeasureText(t, font);
                if (measurement.IsEmpty && ts.FallbackTextMeasurer != null && ts.FallbackTextMeasurer != ts.PrimaryTextMeasurer)
                {
                    measurer = ts.FallbackTextMeasurer;
                    measurement = measurer.MeasureText(t, font);
                }
                if (measurement.IsEmpty && _fontWidthDefault != null)
                {
                    measurement = MeasureGeneric(t, ts, font);
                }
                if (!measurement.IsEmpty && ts.AutofitScaleFactor != 1f)
                {
                    measurement.Height = measurement.Height * ts.AutofitScaleFactor;
                    measurement.Width = measurement.Width * ts.AutofitScaleFactor;
                }
                measureCache.Add(key, measurement);
            }
            return measurement;
        }

        internal TextMeasurement MeasureGeneric(string t, ExcelTextSettings ts, MeasurementFont font)
        {
            TextMeasurement measurement;
            if (FontSize.FontWidths.ContainsKey(font.FontFamily))
            {
                var width = FontSize.GetWidthPixels(font.FontFamily, font.Size);
                var height = FontSize.GetHeightPixels(font.FontFamily, font.Size);
                var defaultWidth = FontSize.GetWidthPixels(FontSize.NonExistingFont, font.Size);
                var defaultHeight = FontSize.GetHeightPixels(FontSize.NonExistingFont, font.Size);
                _nonExistingFont.Size = font.Size;
                _nonExistingFont.Style = font.Style;
                measurement = _genericMeasurer.MeasureText(t, _nonExistingFont);

                measurement.Width *= (float)(width / defaultWidth) * ts.AutofitScaleFactor;
                measurement.Height *= (float)(height / defaultHeight) * ts.AutofitScaleFactor;
            }
            else
            {
                _nonExistingFont.Size = font.Size;
                _nonExistingFont.Style = font.Style;
                measurement = _genericMeasurer.MeasureText(t, _nonExistingFont);
                measurement.Height = measurement.Height * ts.AutofitScaleFactor;
                measurement.Width = measurement.Width * ts.AutofitScaleFactor;
            }

            return measurement;
        }
    }
}
