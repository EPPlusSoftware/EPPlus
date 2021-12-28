using OfficeOpenXml.Core.CellStore;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts;
using OfficeOpenXml.Core.Worksheet.Core.Worksheet.SerializedFonts.Serialization;
using System;
using System.Collections.Generic;
using System.Drawing;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Core
{
    public class AutofitHelper
    {
        private ExcelRangeBase _range;

        public AutofitHelper(ExcelRangeBase range)
        {
            _range = range;
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
            var fontCache = new Dictionary<int, Font>();

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
            if (ws.AutoFilterAddress != null)
            {
                afAddr.Add(new ExcelAddressBase(    ws.AutoFilterAddress._fromRow,
                                                    ws.AutoFilterAddress._fromCol,
                                                    ws.AutoFilterAddress._fromRow,
                                                    ws.AutoFilterAddress._toCol));
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
            var normalXfId = styles.GetNormalStyle().StyleXfId;
            if (normalXfId < 0 || normalXfId >= styles.CellStyleXfs.Count) normalXfId = 0;
            var nf = styles.Fonts[styles.CellStyleXfs[normalXfId].FontId];
            var fs = FontStyle.Regular;
            if (nf.Bold) fs |= FontStyle.Bold;
            if (nf.UnderLine) fs |= FontStyle.Underline;
            if (nf.Italic) fs |= FontStyle.Italic;
            if (nf.Strike) fs |= FontStyle.Strikeout;
            var nfont = new Font(nf.Name, nf.Size, fs);

            var normalSize = Convert.ToSingle(ExcelWorkbook.GetWidthPixels(nf.Name, nf.Size));

            Bitmap b;
            Graphics g;
            float dpiCorrectX, dpiCorrectY;
            try
            {
                //Check for missing GDI+, then use WPF istead.
                b = new Bitmap(1, 1);
                g = Graphics.FromImage(b);
                g.PageUnit = GraphicsUnit.Pixel;
                dpiCorrectX = 96 / g.DpiX;
                dpiCorrectY = 96 / g.DpiY;
            }
            catch
            {
                return;
            }

            #region MeasureString memoization
            // This will call GDI+, and the result isn't cached.
            // Calling this in the tight loop below is very slow.
            var stringFormat = StringFormat.GenericDefault;

            // Sheets usually contain plenty of duplicates
            // Measurestring is very slow, so memoizing yields massive performance benefits.
            // We use the string hash rather than the string to reduce memory load and lookup/compare cost.
            // This means columns can be wrongly calculated on hash collisions. Hash collisions are rare,
            // and they might not affect the size calculation anyway.

            // To support implementations without Tuple/ValueTuple,
            // as well as reduce som overhead, we combine our two
            // 32-bit keys in a single 64-bit value
            var measureCache = new Dictionary<ulong, SizeF>();

            SizeF MeasureString(string t, int fntID)
            {
                ulong key = ((ulong)((uint)t.GetHashCode()) << 32) | (uint)fntID;
                if (!measureCache.TryGetValue(key, out var size))
                {
                    var serializedFontKey = SerializedFontMetrics.GetKey(fontCache[fntID]);
                    if (serializedFontKey < uint.MaxValue)
                    {
                        // uncomment for comparison /MA
                        var s2 = g.MeasureString(t, fontCache[fntID], 10000, stringFormat);
                        var s = TextMeasurer.Measure(t, fontCache[fntID].Size, serializedFontKey);
                        size = new SizeF(s, 0f);
                    }
                    else
                    {
                        size = g.MeasureString(t, fontCache[fntID], 10000, stringFormat);
                    }
                    measureCache.Add(key, size);
                }

                return size;
            }
            #endregion

            foreach (var cell in _range)
            {
                if (ws.Column(cell.Start.Column).Hidden)    //Issue 15338
                    continue;

                if (cell.Merge == true || styles.CellXfs[cell.StyleID].WrapText) continue;
                var fntID = styles.CellXfs[cell.StyleID].FontId;
                Font f;
                if (fontCache.ContainsKey(fntID))
                {
                    f = fontCache[fntID];
                }
                else
                {
                    var fnt = styles.Fonts[fntID];
                    fs = FontStyle.Regular;
                    if (fnt.Bold) fs |= FontStyle.Bold;
                    if (fnt.UnderLine) fs |= FontStyle.Underline;
                    if (fnt.Italic) fs |= FontStyle.Italic;
                    if (fnt.Strike) fs |= FontStyle.Strikeout;
                    f = new Font(fnt.Name, fnt.Size, fs);

                    fontCache.Add(fntID, f);
                }
                var ind = styles.CellXfs[cell.StyleID].Indent;
                var textForWidth = cell.TextForWidth;
                var t = textForWidth + (ind > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', ind) : "");
                if (t.Length > 32000) t = t.Substring(0, 32000); //Issue
                var size = MeasureString(t, fntID);

                double width;
                double r = styles.CellXfs[cell.StyleID].TextRotation;
                if (r <= 0)
                {
                    width = (size.Width * dpiCorrectX + 5) / normalSize;
                }
                else
                {
                    r = (r <= 90 ? r : r - 90);
                    width = (((size.Width * dpiCorrectX - size.Height * dpiCorrectY) * Math.Abs(System.Math.Cos(System.Math.PI * r / 180.0)) + size.Height * dpiCorrectY) + 5) / normalSize;
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
