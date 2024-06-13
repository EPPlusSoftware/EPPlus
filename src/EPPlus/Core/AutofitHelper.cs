/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/01/2022         EPPlus Software AB       EPPlus 6
 *************************************************************************************************/

using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core
{
    internal class AutofitHelper
    {
        private ExcelRangeBase _range;
        ITextMeasurer _genericMeasurer = new GenericFontMetricsTextMeasurer();
        MeasurementFont _nonExistingFont = new MeasurementFont() { FontFamily = FontSize.NonExistingFont };
        Dictionary<float, short> _fontWidthDefault=null;
        Dictionary<int, MeasurementFont> _fontCache;
        ExcelTextSettings _textSettings;
        Dictionary<ulong, TextMeasurement> measureCache = new Dictionary<ulong, TextMeasurement>();
        public AutofitHelper(ExcelRangeBase range)
        {
            _range = range;
            _textSettings = _range._workbook._package.Settings.TextSettings;
            if (FontSize.FontWidths.ContainsKey(FontSize.NonExistingFont))
            {
                FontSize.LoadAllFontsFromResource();
                _fontWidthDefault = FontSize.FontWidths[FontSize.NonExistingFont];
            }
        }

        internal void AutofitColumn(double MinimumWidth, double MaximumWidth)
        {
            var worksheet = _range._worksheet;
            if (worksheet.Dimension == null)
            {
                return;
            }
            if (_range._fromCol < 1 || _range._fromRow < 1)
            {
                _range.SetToSelectedRange();
            }
            var fromCol = _range._fromCol > worksheet.Dimension._fromCol ? _range._fromCol : worksheet.Dimension._fromCol;
            var toCol = _range._toCol < worksheet.Dimension._toCol ? _range._toCol : worksheet.Dimension._toCol;
            var fromRow = _range._fromRow > worksheet.Dimension._fromRow ? _range._fromRow : worksheet.Dimension._fromRow;
            var toRow = _textSettings.AutofitRows > 0 && _textSettings.AutofitRows < _range._toRow ? _textSettings.AutofitRows : _range._toRow;
            toRow = toRow < worksheet.Dimension._toRow ? toRow : worksheet.Dimension._toRow;
            if (fromCol > toCol) return; //Issue 15383
            if (MinimumWidth < 0d)
            {
                MinimumWidth = 0d;
            }
            if (MaximumWidth > 255d)
            {
                MaximumWidth = 255d;
            }
            if(MinimumWidth >= MaximumWidth)
            {
                MinimumWidth = MaximumWidth;
            }

            bool doAdjust = worksheet._package.DoAdjustDrawings;
            worksheet._package.DoAdjustDrawings = false;
            var drawWidths = worksheet.Drawings.GetDrawingWidths();

            _fontCache = new Dictionary<int, MeasurementFont>();
            //Get the font, size and style of the default font
            var styles = worksheet.Workbook.Styles;
            var normalStyle = styles.GetNormalStyle();
            var normalXfId = normalStyle?.StyleXfId ?? 0;
            if (normalXfId < 0 || normalXfId >= styles.CellStyleXfs.Count) normalXfId = 0;
            var normalFont = styles.Fonts[styles.CellStyleXfs[normalXfId].FontId];
            var fontStyle = MeasurementFontStyles.Regular;
            if (normalFont.Bold) fontStyle |= MeasurementFontStyles.Bold;
            if (normalFont.UnderLine) fontStyle |= MeasurementFontStyles.Underline;
            if (normalFont.Italic) fontStyle |= MeasurementFontStyles.Italic;
            if (normalFont.Strike) fontStyle |= MeasurementFontStyles.Strikeout;
            var normalSize = Convert.ToSingle(FontSize.GetWidthPixels(normalFont.Name, normalFont.Size));

            //Get any autofilter to widen these columns
            var afAddr = new List<ExcelAddressBase>();
            if (worksheet.AutoFilter.Address != null)
            {
                afAddr.Add(new ExcelAddressBase(worksheet.AutoFilter.Address._fromRow,
                                                    worksheet.AutoFilter.Address._fromCol,
                                                    worksheet.AutoFilter.Address._fromRow,
                                                    worksheet.AutoFilter.Address._toCol));
                afAddr[afAddr.Count - 1]._ws = _range.WorkSheetName;
            }
            foreach (var tbl in worksheet.Tables)
            {
                if (tbl.AutoFilterAddress != null)
                {
                    afAddr.Add(new ExcelAddressBase(tbl.AutoFilterAddress._fromRow,
                                                                            tbl.AutoFilterAddress._fromCol,
                                                                            tbl.AutoFilterAddress._fromRow,
                                                                            tbl.AutoFilterAddress._toCol));
                    afAddr[afAddr.Count - 1]._ws = _range.WorkSheetName;
                }
            }
            for (int col = fromCol; col <= toCol; col++)
            {
                if (worksheet.Column(col).Hidden)    //Issue 15338
                {
                    continue;
                }
                if (worksheet.Column(col).Width >= MaximumWidth)
                {
                    continue;
                }
                var currentMaxWidth = 0d;
                Dictionary<MeasurementFont, int> textLengthCache = new Dictionary<MeasurementFont, int>();
                foreach (var af in afAddr)
                {
                    if (af.Collide(fromRow, col, toRow, col) != eAddressCollition.No)
                    {
                        var cell = worksheet.Cells[af.Address];
                        var cellStyleId = styles.CellXfs[cell.StyleID];
                        currentMaxWidth = GetTextLength(cell, textLengthCache, styles, cellStyleId, normalSize, MaximumWidth, currentMaxWidth);
                    }
                }
                foreach (var cell in worksheet.Cells[fromRow, col, toRow, col])
                {
                    var cellStyleId = styles.CellXfs[cell.StyleID];
                    if (cell.Merge == true || cellStyleId.WrapText) continue;
                    currentMaxWidth = GetTextLength(cell, textLengthCache, styles, cellStyleId, normalSize, MaximumWidth, currentMaxWidth);
                    if(currentMaxWidth >= MaximumWidth)
                    {
                        break;
                    }
                }
                if (currentMaxWidth < MinimumWidth)
                {
                    currentMaxWidth = MinimumWidth;
                }
                worksheet.Column(col).Width = currentMaxWidth;
            }
            worksheet.Drawings.AdjustWidth(drawWidths);
            worksheet._package.DoAdjustDrawings = doAdjust;
        }

        private double GetTextLength(ExcelRangeBase cell, Dictionary<MeasurementFont, int> textLengthCache, ExcelStyles styles, Style.XmlAccess.ExcelXfs cellStyleId, float normalSize, double MaximumWidth, double currentMaxWidth)
        {
            var fontID = cellStyleId.FontId;
            MeasurementFont measurementFont;
            if (_fontCache.ContainsKey(fontID))
            {
                measurementFont = _fontCache[fontID];
            }
            else
            {
                var font = styles.Fonts[fontID];
                var fontStyle = MeasurementFontStyles.Regular;
                if (font.Bold) fontStyle |= MeasurementFontStyles.Bold;
                if (font.UnderLine) fontStyle |= MeasurementFontStyles.Underline;
                if (font.Italic) fontStyle |= MeasurementFontStyles.Italic;
                if (font.Strike) fontStyle |= MeasurementFontStyles.Strikeout;
                measurementFont = new MeasurementFont
                {
                    FontFamily = font.Name,
                    Style = fontStyle,
                    Size = font.Size
                };
                _fontCache.Add(fontID, measurementFont);
            }
            var indent = cellStyleId.Indent;
            var textForWidth = cell.TextForWidth;
            var text = textForWidth + (indent > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', indent) : "");
            if (text.Length > 32000) { text = text.Substring(0, 32000); } //Issue

            if(textLengthCache.ContainsKey(measurementFont) && text.Length < textLengthCache[measurementFont] * _textSettings.textLengthThreshold)
            {
                return currentMaxWidth;
            }

            var size = MeasureString(text, fontID, measureCache);

            double width;
            double rotation = cellStyleId.TextRotation;
            if (rotation <= 0)
            {
                var padding = 0; // 5
                width = (size.Width + padding) / normalSize;
            }
            else
            {
                rotation = (rotation <= 90 ? rotation : rotation - 90);
                width = (((size.Width - size.Height) * Math.Abs(System.Math.Cos(System.Math.PI * rotation / 180.0)) + size.Height) + 5) / normalSize;
            }
            if (currentMaxWidth < width)
            {
                currentMaxWidth = width;
                if (!textLengthCache.ContainsKey(measurementFont))
                {
                    textLengthCache.Add(measurementFont, text.Length);
                }
                else
                {
                    textLengthCache[measurementFont] = text.Length;
                }
            }
            if (currentMaxWidth >= MaximumWidth)
            {
                currentMaxWidth = MaximumWidth;
            }
            return currentMaxWidth;
        }

        private TextMeasurement MeasureString(string text, int fontID, Dictionary<ulong, TextMeasurement> measureCache)
        {
            ulong key = ((ulong)((uint)text.GetHashCode()) << 32) | (uint)fontID;
            if (!measureCache.TryGetValue(key, out var measurement))
            {
                var measurer = _textSettings.PrimaryTextMeasurer;
                var font = _fontCache[fontID];
                measurement = measurer.MeasureText(text, font);
                if (measurement.IsEmpty && _textSettings.FallbackTextMeasurer != null && _textSettings.FallbackTextMeasurer != _textSettings.PrimaryTextMeasurer)
                {
                    measurer = _textSettings.FallbackTextMeasurer;
                    measurement = measurer.MeasureText(text, font);
                }
                if (measurement.IsEmpty && _fontWidthDefault != null)
                {
                    measurement = MeasureGeneric(text, _textSettings, font);
                }
                if (!measurement.IsEmpty && _textSettings.AutofitScaleFactor != 1f)
                {
                    measurement.Height = measurement.Height * _textSettings.AutofitScaleFactor;
                    measurement.Width = measurement.Width * _textSettings.AutofitScaleFactor;
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
