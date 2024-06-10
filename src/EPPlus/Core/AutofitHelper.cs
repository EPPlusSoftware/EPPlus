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
using OfficeOpenXml.Core.RangeQuadTree;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.Fonts.GenericMeasurements;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
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
        public AutofitHelper(ExcelRangeBase range)
        {
            _range = range;
            if(FontSize.FontWidths.ContainsKey(FontSize.NonExistingFont))
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
            var fromCol = _range._fromCol;
            var toCol = _range._toCol;
            var fromRow = _range._fromRow;
            var toRow = _range._toRow;
            if (fromCol > toCol) return; //Issue 15383
            if (_range.Addresses == null)
            {
                SetMinWidth(worksheet, MinimumWidth, fromCol, toCol);
            }
            else
            {
                foreach (var addr in _range.Addresses)
                {
                    fromCol = addr._fromCol > worksheet.Dimension._fromCol ? addr._fromCol : worksheet.Dimension._fromCol;
                    toCol = addr._toCol < worksheet.Dimension._toCol ? addr._toCol : worksheet.Dimension._toCol;
                    SetMinWidth(worksheet, MinimumWidth, fromCol, toCol);
                }
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
            var textSettings = _range._workbook._package.Settings.TextSettings;

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
                foreach (var af in afAddr)
                {
                    if (af.Collide(_range._fromRow, col, _range._toRow, col) != eAddressCollition.No)
                    {
                        var cell = _range.Worksheet.Cells[af.Address];
                        currentMaxWidth = GetTextLength(cell, textSettings, styles, normalSize, MaximumWidth, currentMaxWidth);
                    }
                }
                foreach (var cell in _range.Worksheet.Cells[fromRow, col, toRow, col])
                {
                    if (cell.Merge == true || styles.CellXfs[cell.StyleID].WrapText) continue;
                    currentMaxWidth = GetTextLength(cell, textSettings, styles, normalSize, MaximumWidth, currentMaxWidth);
                }
                if (currentMaxWidth < MinimumWidth)
                {
                    currentMaxWidth = MinimumWidth;
                }
                worksheet.Column(col).Width = currentMaxWidth;
            }
            worksheet.Drawings.AdjustWidth(drawWidths);
            worksheet._package.DoAdjustDrawings = doAdjust;

            //if (worksheet.Dimension == null)
            //{
            //    return;
            //}
            //if (_range._fromCol < 1 || _range._fromRow < 1)
            //{
            //    _range.SetToSelectedRange();
            //}
            //_fontCache = new Dictionary<int, MeasurementFont>();

            //bool doAdjust = worksheet._package.DoAdjustDrawings;
            //worksheet._package.DoAdjustDrawings = false;
            //var drawWidths = worksheet.Drawings.GetDrawingWidths();

            ////set up fromCol & toCol
            //var fromCol = _range._fromCol > worksheet.Dimension._fromCol ? _range._fromCol : worksheet.Dimension._fromCol;
            //var toCol = _range._toCol < worksheet.Dimension._toCol ? _range._toCol : worksheet.Dimension._toCol;

            //if (fromCol > toCol) return; //Issue 15383

            //if (_range.Addresses == null)
            //{
            //    SetMinWidth(worksheet, MinimumWidth, fromCol, toCol);
            //}
            //else
            //{
            //    foreach (var addr in _range.Addresses)
            //    {
            //        fromCol = addr._fromCol > worksheet.Dimension._fromCol ? addr._fromCol : worksheet.Dimension._fromCol;
            //        toCol = addr._toCol < worksheet.Dimension._toCol ? addr._toCol : worksheet.Dimension._toCol;
            //        SetMinWidth(worksheet, MinimumWidth, fromCol, toCol);
            //    }
            //}

            ////Get any autofilter to widen these columns
            //var afAddr = new List<ExcelAddressBase>();
            //if (worksheet.AutoFilter.Address != null)
            //{
            //    afAddr.Add(new ExcelAddressBase(    worksheet.AutoFilter.Address._fromRow,
            //                                        worksheet.AutoFilter.Address._fromCol,
            //                                        worksheet.AutoFilter.Address._fromRow,
            //                                        worksheet.AutoFilter.Address._toCol));
            //    afAddr[afAddr.Count - 1]._ws = _range.WorkSheetName;
            //}
            //foreach (var tbl in worksheet.Tables)
            //{
            //    if (tbl.AutoFilterAddress != null)
            //    {
            //        afAddr.Add(new ExcelAddressBase(tbl.AutoFilterAddress._fromRow,
            //                                                                tbl.AutoFilterAddress._fromCol,
            //                                                                tbl.AutoFilterAddress._fromRow,
            //                                                                tbl.AutoFilterAddress._toCol));
            //        afAddr[afAddr.Count - 1]._ws = _range.WorkSheetName;
            //    }
            //}

            ////Get the font, size and style of the default font
            //var styles = worksheet.Workbook.Styles;
            //var normalStyle = styles.GetNormalStyle();
            //var normalXfId = normalStyle?.StyleXfId ?? 0;
            //if (normalXfId < 0 || normalXfId >= styles.CellStyleXfs.Count) normalXfId = 0;
            //var normalFont = styles.Fonts[styles.CellStyleXfs[normalXfId].FontId];
            //var fontStyle = MeasurementFontStyles.Regular;
            //if (normalFont.Bold) fontStyle |= MeasurementFontStyles.Bold;
            //if (normalFont.UnderLine) fontStyle |= MeasurementFontStyles.Underline;
            //if (normalFont.Italic) fontStyle |= MeasurementFontStyles.Italic;
            //if (normalFont.Strike) fontStyle |= MeasurementFontStyles.Strikeout;
            //var nfont = new MeasurementFont
            //{
            //    FontFamily = normalFont.Name,
            //    Style = fontStyle,
            //    Size = normalFont.Size
            //};

            //var normalSize = Convert.ToSingle(FontSize.GetWidthPixels(normalFont.Name, normalFont.Size));
            //var textSettings = _range._workbook._package.Settings.TextSettings;
            //Go through all cells
            //foreach (var cell in _range)
            //{
            //if (worksheet.Column(cell.Start.Column).Hidden)    //Issue 15338
            //{
            //    continue;
            //}
            //if (worksheet.Column(cell._fromCol).Width >= MaximumWidth)
            //{
            //    continue;
            //}

            //if (cell.Merge == true || styles.CellXfs[cell.StyleID].WrapText) continue;
            //var fontID = styles.CellXfs[cell.StyleID].FontId;
            //MeasurementFont measurementFont;
            //if (_fontCache.ContainsKey(fontID))
            //{
            //    measurementFont = _fontCache[fontID];
            //}
            //else
            //{
            //    var font = styles.Fonts[fontID];
            //    fontStyle = MeasurementFontStyles.Regular;
            //    if (font.Bold) fontStyle |= MeasurementFontStyles.Bold;
            //    if (font.UnderLine) fontStyle |= MeasurementFontStyles.Underline;
            //    if (font.Italic) fontStyle |= MeasurementFontStyles.Italic;
            //    if (font.Strike) fontStyle |= MeasurementFontStyles.Strikeout;
            //    measurementFont = new MeasurementFont
            //    {
            //        FontFamily = font.Name,
            //        Style = fontStyle,
            //        Size = font.Size
            //    };
            //    _fontCache.Add(fontID, measurementFont);
            //}
            //var indent = styles.CellXfs[cell.StyleID].Indent;
            //var textForWidth = cell.TextForWidth;
            //var text = textForWidth + (indent > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', indent) : "");
            //if (text.Length > 32000) text = text.Substring(0, 32000); //Issue
            //var size = MeasureString(text, fontID, textSettings);

            //double width;
            //double rotation = styles.CellXfs[cell.StyleID].TextRotation;
            //if (rotation <= 0)
            //{
            //    var padding = 0; // 5
            //    width = (size.Width + padding) / normalSize;
            //}
            //else
            //{
            //    rotation = (rotation <= 90 ? rotation : rotation - 90);
            //    width = (((size.Width - size.Height) * Math.Abs(System.Math.Cos(System.Math.PI * rotation / 180.0)) + size.Height) + 5) / normalSize;
            //}

            //foreach (var a in afAddr)
            //{
            //    if (a.Collide(cell) != eAddressCollition.No)
            //    {
            //        width += 2.25;
            //        break;
            //    }
            //}

            //if (width > worksheet.Column(cell._fromCol).Width )
            //{
            //    worksheet.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
            //}
            //}
            //worksheet.Drawings.AdjustWidth(drawWidths);
            //worksheet._package.DoAdjustDrawings = doAdjust;
        }

        private double GetTextLength(ExcelRangeBase cell, ExcelTextSettings textSettings, ExcelStyles styles, float normalSize, double MaximumWidth, double currentMaxWidth)
        {
            var fontID = styles.CellXfs[cell.StyleID].FontId;
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
            var indent = styles.CellXfs[cell.StyleID].Indent;
            var textForWidth = cell.TextForWidth;
            var text = textForWidth + (indent > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', indent) : "");
            if (text.Length > 32000) { text = text.Substring(0, 32000); } //Issue
            var size = MeasureString(text, fontID, textSettings);

            double width;
            double rotation = styles.CellXfs[cell.StyleID].TextRotation;
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
            }
            if (currentMaxWidth >= MaximumWidth)
            {
                currentMaxWidth = MaximumWidth;
            }
            return currentMaxWidth;
        }

        private TextMeasurement MeasureString(string text, int fontID, ExcelTextSettings textSettings)
        {
            var measureCache = new Dictionary<ulong, TextMeasurement>();
            ulong key = ((ulong)((uint)text.GetHashCode()) << 32) | (uint)fontID;
            if (!measureCache.TryGetValue(key, out var measurement))
            {
                var measurer = textSettings.PrimaryTextMeasurer;
                var font = _fontCache[fontID];
                measurement = measurer.MeasureText(text, font);
                if (measurement.IsEmpty && textSettings.FallbackTextMeasurer != null && textSettings.FallbackTextMeasurer != textSettings.PrimaryTextMeasurer)
                {
                    measurer = textSettings.FallbackTextMeasurer;
                    measurement = measurer.MeasureText(text, font);
                }
                if (measurement.IsEmpty && _fontWidthDefault != null)
                {
                    measurement = MeasureGeneric(text, textSettings, font);
                }
                if (!measurement.IsEmpty && textSettings.AutofitScaleFactor != 1f)
                {
                    measurement.Height = measurement.Height * textSettings.AutofitScaleFactor;
                    measurement.Width = measurement.Width * textSettings.AutofitScaleFactor;
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

        private void SetMinWidth(ExcelWorksheet worksheet, double minimumWidth, int fromCol, int toCol)
        {
            var iterator = new CellStoreEnumerator<ExcelValue>(worksheet._values, 0, fromCol, 0, toCol);
            var prevCol = fromCol;
            foreach (ExcelValue excelValue in iterator)
            {
                var col = (ExcelColumn)excelValue._value;
                if (col.Hidden) continue;
                col.Width = minimumWidth;
                if (worksheet.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
                {
                    var newCol = worksheet.Column(prevCol);
                    newCol.ColumnMax = col.ColumnMin - 1;
                    newCol.Width = minimumWidth;
                }
                prevCol = col.ColumnMax + 1;
            }
            if (worksheet.DefaultColWidth > minimumWidth && prevCol < toCol)
            {
                var newCol = worksheet.Column(prevCol);
                newCol.ColumnMax = toCol;
                newCol.Width = minimumWidth;
            }
        }
    }
}
