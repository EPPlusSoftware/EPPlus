using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Core.CellMeasurements
{
    internal class CellWidthMeasurer
    {
        public CellWidthMeasurer(ExcelWorksheet sheet, int startRow, int startCol, int endRow, int endCol)
        {
            _sheet = sheet;
            _startRow = startRow;
            _startCol = startCol;
            _endRow = endRow;
            _endCol = endCol;
        }

        private readonly ExcelWorksheet _sheet;
        private readonly int _startRow;
        private readonly int _endRow;
        private readonly int _startCol;
        private readonly int _endCol;
        ITextMeasurer _genericMeasurer = new GenericFontMetricsTextMeasurer();
        MeasurementFont _nonExistingFont = new MeasurementFont() { FontFamily = FontSize.NonExistingFont };
        Dictionary<float, short> _fontWidthDefault = null;
        Dictionary<int, MeasurementFont> _fontCache;

        public Dictionary<int, CellWidthMeasurement> Measure()
        {
            _fontCache = new Dictionary<int, MeasurementFont>();
            var styles = _sheet.Workbook.Styles;
            var textSettings = _sheet.Workbook._package.Settings.TextSettings;
            var result = new Dictionary<int, CellWidthMeasurement>();
            for (var column = _startCol; column <= _endCol; column++)
            {
                float floatMaxWidth = 0;
                for (var row = _startRow; row <= _endRow; row++)
                {
                    var cell = _sheet.Cells[row, column];
                    if (cell.Value == null) continue;
                    var fntId = styles.CellXfs[cell.StyleID].FontId;
                    var fs = MeasurementFontStyles.Regular;
                    MeasurementFont f;
                    if (_fontCache.ContainsKey(fntId))
                    {
                        f = _fontCache[fntId];
                    }
                    else
                    {
                        var fnt = styles.Fonts[fntId];
                        fs = MeasurementFontStyles.Regular;
                        if (fnt.Bold) fs |= MeasurementFontStyles.Bold;
                        if (fnt.UnderLine) fs |= MeasurementFontStyles.Underline;
                        if (fnt.Italic) fs |= MeasurementFontStyles.Italic;
                        if (fnt.Strike) fs |= MeasurementFontStyles.Strikeout;
                        f = new MeasurementFont
                        {
                            FontFamily = fnt.Name,
                            Style = fs,
                            Size = fnt.Size
                        };

                        _fontCache.Add(fntId, f);
                    }
                    var measurement = MeasureString(cell.Value.ToString(), fntId, textSettings);
                    if (measurement.Width > floatMaxWidth)
                    {
                        floatMaxWidth = measurement.Width;
                        if(!result.ContainsKey(column))
                        {
                            result[column] = new CellWidthMeasurement { MaxWidth = measurement.Width };
                        }
                        else
                        {
                            result[column].MaxWidth = measurement.Width;
                        }
                    }
                }
            }
            return result;
        }

        private TextMeasurement MeasureString(string t, int fntID, ExcelTextSettings ts)
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

        private TextMeasurement MeasureGeneric(string t, ExcelTextSettings ts, MeasurementFont font)
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
