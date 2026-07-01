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

using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core
{
    internal class AutofitHelper
    {
        // Approximate width in pixels (at 96 DPI) of the autofilter dropdown arrow rendered by Excel.
        // Set to 22d (17 pixels physical dropdown button width + 5 pixels extra padding to prevent text from sticking).
        private const double AutoFilterArrowWidthPixels = 22d;
        private ExcelRangeBase _range;
        ITextMeasurer _genericMeasurer = new GenericFontMetricsTextMeasurer();
        MeasurementFont _nonExistingFont = new MeasurementFont() { FontFamily = FontSize.NonExistingFont };
        Dictionary<float, short> _fontWidthDefault = null;
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
            var fromCol = Math.Max(_range._fromCol, worksheet.Dimension._fromCol);
            var toCol = Math.Min(_range._toCol, worksheet.Dimension._toCol);
            var fromRow = _range._fromRow;
            var toRow = _textSettings.AutofitRows > 0 && _textSettings.AutofitRows < _range._toRow ? _textSettings.AutofitRows : _range._toRow;
            if (fromCol > toCol) return; //Issue 15383
            if (MinimumWidth < 0d)
            {
                MinimumWidth = 0d;
            }
            if (MaximumWidth > 265d)
            {
                MaximumWidth = 256d;
            }
            if (MinimumWidth >= MaximumWidth)
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

            //Get any auto filter to widen these columns
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
                        var cell = worksheet.Cells[af._fromRow, col];
                        var cellStyleId = styles.CellXfs[cell.StyleID];
                        if (cellStyleId.WrapText && _textSettings.MeasureWrappedTextCells == false) continue;

                        // Reserve room for the autofilter dropdown arrow. The arrow is a fixed-size UI
                        // element (17px at 96 DPI + 5px padding). Excel renders this dropdown arrow *only* on the last
                        // line of a header cell. Preceding lines in a wrapped cell can overlap into the
                        // dropdown area. To avoid inflating the column width unnecessarily, we pass this
                        // 22px padding as `lastLinePadding` so that the text measurer only adds it to
                        // the last wrapped line of the cell.
                        float lastLinePadding = (float)AutoFilterArrowWidthPixels;
                        currentMaxWidth = GetTextLength(cell, textLengthCache, styles, cellStyleId, normalSize, MaximumWidth, currentMaxWidth, lastLinePadding);
                        if (currentMaxWidth >= MaximumWidth)
                        {
                            currentMaxWidth = MaximumWidth;
                        }
                    }
                }
                foreach (var cell in worksheet.Cells[fromRow, col, toRow, col])
                {
                    var cellStyleId = styles.CellXfs[cell.StyleID];
                    if (cell.Merge == true) continue;
                    if (cellStyleId.WrapText && _textSettings.MeasureWrappedTextCells == false) continue;
                    currentMaxWidth = GetTextLength(cell, textLengthCache, styles, cellStyleId, normalSize, MaximumWidth, currentMaxWidth);
                    if (currentMaxWidth >= MaximumWidth)
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

        /// <summary>
        /// Calculates the text width of a cell and updates the current maximum column width
        /// </summary>
        /// <param name="cell">The cell to measure</param>
        /// <param name="textLengthCache">Cache for measured font lengths</param>
        /// <param name="styles">The Excel styles collection</param>
        /// <param name="cellStyleId">The XML style definition of the cell</param>
        /// <param name="normalSize">The width of the normal font's reference char in pixels</param>
        /// <param name="MaximumWidth">The maximum allowed column width</param>
        /// <param name="currentMaxWidth">The currently tracked maximum width of the column</param>
        /// <param name="lastLinePadding">Optional padding in pixels to apply only to the last wrapped line of text (e.g. for AutoFilter dropdown arrow). Defaults to 0f because most cells do not have an adjacent dropdown.</param>
        /// <returns>The updated maximum column width</returns>
        private double GetTextLength(ExcelRangeBase cell, Dictionary<MeasurementFont, int> textLengthCache, ExcelStyles styles, Style.XmlAccess.ExcelXfs cellStyleId, float normalSize, double MaximumWidth, double currentMaxWidth, float lastLinePadding = 0f)
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

            if(cell.Style.WrapText==false)
            {
                text = text.Replace("\r","").Replace("\n","");
            }

            if (textLengthCache.ContainsKey(measurementFont) && text.Length < textLengthCache[measurementFont] * _textSettings.textLengthThreshold)
            {
                return currentMaxWidth;
            }
            var size = MeasureString(text, fontID, cellStyleId.WrapText, lastLinePadding, measureCache);

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

        /// <summary>
        /// Measures a text string using the specified font, with caching support
        /// </summary>
        /// <param name="text">The text to measure</param>
        /// <param name="fontID">ID of the font to use</param>
        /// <param name="wrapText">Whether word wrapping is enabled on the cell style</param>
        /// <param name="lastLinePadding">Padding in pixels to apply only to the last line's width</param>
        /// <param name="measureCache">High-performance lookup cache for measured strings</param>
        /// <returns>A <see cref="TextMeasurement"/> representing the width and height</returns>
        private TextMeasurement MeasureString(string text, int fontID, bool wrapText, float lastLinePadding, Dictionary<ulong, TextMeasurement> measureCache)
        {
            // Create a high-performance 64-bit cache key based on the original key calculation.
            // Original baseline:
            //   ulong key = ((ulong)((uint)text.GetHashCode()) << 32) | (uint)fontID;
            //
            // We extend this by packing wrapText and lastLinePadding flags into the high bits of the lower 32-bit word.

            // 1. Shift the 32-bit string hash code into the upper 32 bits (bits 32-63)
            ulong hashPart = (ulong)(uint)text.GetHashCode() << 32;

            // 2. Set boolean flags in the upper two bits of the lower 32-bit word (bits 30 and 31)
            ulong paddingFlag = lastLinePadding > 0f ? (1UL << 31) : 0UL;
            ulong wrapFlag = wrapText ? (1UL << 30) : 0UL;

            // 3. Mask the font ID to the remaining 30 bits (bits 0-29).
            // This supports up to 1,073,741,823 unique fonts while preventing any bit-flag overflow.
            ulong fontPart = (ulong)fontID & 0x3FFFFFFF;

            // 4. Combine all parts into a single 64-bit key
            ulong key = hashPart | paddingFlag | wrapFlag | fontPart;

            if (!measureCache.TryGetValue(key, out var measurement))
            {
                var font = _fontCache[fontID];

                measurement = _textSettings.PrimaryTextMeasurer.MeasureText(text, font, wrapText, lastLinePadding);

                if (measurement.IsEmpty && _textSettings.FallbackTextMeasurer != null && _textSettings.FallbackTextMeasurer != _textSettings.PrimaryTextMeasurer)
                {
                    measurement = _textSettings.FallbackTextMeasurer.MeasureText(text, font, wrapText, lastLinePadding);
                }
                if (measurement.IsEmpty && _fontWidthDefault != null)
                {
                    measurement = MeasureGeneric(text, _textSettings, font);
                    measurement.Width += lastLinePadding;
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