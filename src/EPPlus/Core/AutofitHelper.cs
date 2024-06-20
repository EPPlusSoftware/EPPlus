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
                _fontWidthDefault = FontSize.FontWidths[FontSize.NonExistingFont];            }

        }

        internal void AutofitColumn(double MinimumWidth, double MaximumWidth)
        {
            var ws = _range._worksheet;
            if (ws.Dimension == null)
            {
                return;
            }
            if (_range._fromCol < 1 || _range._fromRow < 1)
            {
                _range.SetToSelectedRange();
            }
            _fontCache = new Dictionary<int, MeasurementFont>();

            bool doAdjust = ws._package.DoAdjustDrawings;
            ws._package.DoAdjustDrawings = false;
            var drawWidths = ws.Drawings.GetDrawingWidths();

            var fromCol = _range._fromCol > ws.Dimension._fromCol ? _range._fromCol : ws.Dimension._fromCol;
            var toCol = _range._toCol < ws.Dimension._toCol ? _range._toCol : ws.Dimension._toCol;

            if (fromCol > toCol) return; //Issue 15383

            if (_range.Addresses == null)
            {
                SetMinWidth(ws, MinimumWidth, fromCol, toCol);
            }
            else
            {
                foreach (var addr in _range.Addresses)
                {
                    fromCol = addr._fromCol > ws.Dimension._fromCol ? addr._fromCol : ws.Dimension._fromCol;
                    toCol = addr._toCol < ws.Dimension._toCol ? addr._toCol : ws.Dimension._toCol;
                    SetMinWidth(ws, MinimumWidth, fromCol, toCol);
                }
            }

            //Get any autofilter to widen these columns
            var afAddr = new List<ExcelAddressBase>();
            if (ws.AutoFilter.Address != null)
            {
                afAddr.Add(new ExcelAddressBase(    ws.AutoFilter.Address._fromRow,
                                                    ws.AutoFilter.Address._fromCol,
                                                    ws.AutoFilter.Address._fromRow,
                                                    ws.AutoFilter.Address._toCol));
                afAddr[afAddr.Count - 1]._ws = _range.WorkSheetName;
            }
            foreach (var tbl in ws.Tables)
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

            var styles = ws.Workbook.Styles;
            var ns = styles.GetNormalStyle();
            var normalXfId = ns?.StyleXfId ?? 0;
            if (normalXfId < 0 || normalXfId >= styles.CellStyleXfs.Count) normalXfId = 0;
            var nf = styles.Fonts[styles.CellStyleXfs[normalXfId].FontId];
            var fs = MeasurementFontStyles.Regular;
            if (nf.Bold) fs |= MeasurementFontStyles.Bold;
            if (nf.UnderLine) fs |= MeasurementFontStyles.Underline;
            if (nf.Italic) fs |= MeasurementFontStyles.Italic;
            if (nf.Strike) fs |= MeasurementFontStyles.Strikeout;
            var nfont = new MeasurementFont
            {
                FontFamily = nf.Name,
                Style = fs,
                Size = nf.Size
            };

            var normalSize = Convert.ToSingle(FontSize.GetWidthPixels(nf.Name, nf.Size));
            var textSettings = _range._workbook._package.Settings.TextSettings;

            foreach (var cell in _range)
            {
                if (ws.Column(cell.Start.Column).Hidden)    //Issue 15338
                    continue;

                if (cell.Merge == true || styles.CellXfs[cell.StyleID].WrapText) continue;
                var fntID = styles.CellXfs[cell.StyleID].FontId;
                MeasurementFont f;
                if (_fontCache.ContainsKey(fntID))
                {
                    f = _fontCache[fntID];
                }
                else
                {
                    var fnt = styles.Fonts[fntID];
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

                    _fontCache.Add(fntID, f);
                }
                var ind = styles.CellXfs[cell.StyleID].Indent;
                var textForWidth = cell.TextForWidth;
                var t = textForWidth + (ind > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', ind) : "");
                if (t.Length > 32000) t = t.Substring(0, 32000); //Issue
                var size = MeasureString(t, fntID, textSettings);

                double width;
                double r = styles.CellXfs[cell.StyleID].TextRotation;
                if (r <= 0)
                {
                    var padding = 0; // 5
                    width = (size.Width + padding) / normalSize;
                }
                else
                {
                    r = (r <= 90 ? r : r - 90);
                    width = (((size.Width - size.Height) * Math.Abs(System.Math.Cos(System.Math.PI * r / 180.0)) + size.Height) + 5) / normalSize;
                }

                foreach (var a in afAddr)
                {
                    if (a.Collide(cell) != eAddressCollition.No)
                    {
                        width += 2.25;
                        break;
                    }
                }

                if (width > ws.Column(cell._fromCol).Width)
                {
                    ws.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
                }
            }
            ws.Drawings.AdjustWidth(drawWidths);
            ws._package.DoAdjustDrawings = doAdjust;
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

        private void SetMinWidth(ExcelWorksheet ws, double minimumWidth, int fromCol, int toCol)
        {
            var iterator = new CellStoreEnumerator<ExcelValue>(ws._values, 0, fromCol, 0, toCol);
            var prevCol = fromCol;
            foreach (ExcelValue val in iterator)
            {
                var col = (ExcelColumn)val._value;
                if (col.Hidden) continue;
                col.Width = minimumWidth;
                if (ws.DefaultColWidth > minimumWidth && col.ColumnMin > prevCol)
                {
                    var newCol = ws.Column(prevCol);
                    newCol.ColumnMax = col.ColumnMin - 1;
                    newCol.Width = minimumWidth;
                }
                prevCol = col.ColumnMax + 1;
            }
            if (ws.DefaultColWidth > minimumWidth && prevCol < toCol)
            {
                var newCol = ws.Column(prevCol);
                newCol.ColumnMax = toCol;
                newCol.Width = minimumWidth;
            }
        }
    }
}
