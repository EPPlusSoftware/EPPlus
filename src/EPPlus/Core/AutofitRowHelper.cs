/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  20/12/2022         EPPlus Software AB       EPPlus 6
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Core
{
    internal class AutofitRowHelper
    {
        private ExcelRangeBase _range;
        private readonly TextMeasureUtility _textMeasureUtility = new TextMeasureUtility();
        public AutofitRowHelper(ExcelRangeBase range)
        {
            _range = range;
        }

        private double GetWrappedTextHeight(string txt, int fntId, double columnWidth, ExcelTextSettings textSettings, float normalSize)
        {
            var paragraphs = txt.Split(new[] { '\n', '\r'  }, StringSplitOptions.RemoveEmptyEntries);
            var txtArr = txt.Split(' ');
            var rows = new List<StringBuilder>
            {
                new StringBuilder()
            };
            var rowIx = 0;
            foreach(var word in txtArr)
            {
                if (rows[0].Length > 0)
                {
                    var testWord = rows[0].ToString() + " " + word;
                    var measurement = _textMeasureUtility.MeasureString(testWord, fntId, textSettings);
                    var width = measurement.Width / normalSize;
                    if(width > columnWidth)
                    {
                        rows.Add(new StringBuilder());
                        rowIx++;
                    }
                    rows[rowIx].Append(" " + word);
                }
                else
                {
                    rows[rowIx].Append(word);
                }
            }
            var res = new StringBuilder();
            foreach(var row in rows)
            {
                res.Append(row.ToString());
                res.Append("\n");
            }
            var m = _textMeasureUtility.MeasureString(res.ToString(), fntId, textSettings);
            var height = m.Height * 0.75d;
            height = System.Math.Round(height, 1);
            return height;
        }

        internal void AutofitRows()
        {
            var ws = _range._worksheet;
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

            var rowIx = _range.Start.Row;
            var startRow = _range.Start.Row;
            var endRow = _range.End.Row;
            while(rowIx <= endRow)
            {
                if (ws.Row(rowIx).Hidden)
                {
                    rowIx++;
                    continue;
                }

                var rowRange = ws.Cells[_range.Start.Column, _range.End.Column, rowIx, rowIx];

                var startCol = _range.Start.Column;
                var endCol = _range.End.Column;
                for(var colIx = startCol; colIx <= endCol; colIx++)
                {
                    var cell = ws.Cells[rowIx, colIx];
                    if (!cell.Style.WrapText) continue;
                    var colWidth = ws.Column(colIx).Width;
                    //if (cell.Merge == true || styles.CellXfs[cell.StyleID].WrapText) continue;
                    var fntID = styles.CellXfs[cell.StyleID].FontId;
                    MeasurementFont f;
                    if (_textMeasureUtility.FontCache.ContainsKey(fntID))
                    {
                        f = _textMeasureUtility.FontCache[fntID];
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

                        _textMeasureUtility.FontCache.Add(fntID, f);
                    }
                    var h = GetWrappedTextHeight(cell.Text, fntID, colWidth, textSettings, normalSize);
                    if(h > ws.Row(rowIx).Height)
                    {
                        ws.Row(rowIx).Height = h;
                    }
                }
                rowIx++;
            }
        }
    }
}
