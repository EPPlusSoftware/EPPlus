using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using static OfficeOpenXml.ExcelAddressBase;

namespace OfficeOpenXml.Utils.AutofitCols
{
    internal class MeasurementTask
    {
        public MeasurementTask(IEnumerable<ExcelRangeBase> cells, FontCache fontCache, List<ExcelAddressBase> autofilterAddresses)
        {
            _cells = cells;
            var firstCell = _cells.FirstOrDefault();
            if(firstCell != null)
                _worksheet = firstCell._worksheet;
            _fontCache = fontCache;
            _afAddr = autofilterAddresses;
        }

        private readonly IEnumerable<ExcelRangeBase> _cells;
        private readonly ExcelWorksheet _worksheet;
        private readonly FontCache _fontCache;
        private readonly List<ExcelAddressBase> _afAddr;
        private readonly ColumnWidthMap _maxWidthMap = new ColumnWidthMap();

        internal IDictionary<int, double> GetResult()
        {
            return _maxWidthMap.GetResult();
        }

        internal void Execute()
        {
            if (_worksheet == null) return;
            var styles = _worksheet.Workbook.Styles;
            var nf = styles.Fonts[styles.CellXfs[0].FontId];
            var fs = FontStyle.Regular;
            if (nf.Bold) fs |= FontStyle.Bold;
            if (nf.UnderLine) fs |= FontStyle.Underline;
            if (nf.Italic) fs |= FontStyle.Italic;
            if (nf.Strike) fs |= FontStyle.Strikeout;
            var nfont = new Font(nf.Name, nf.Size, fs);

            var normalSize = Convert.ToSingle(ExcelWorkbook.GetWidthPixels(nf.Name, nf.Size));

            Bitmap b;
            Graphics g = null;
            try
            {
                //Check for missing GDI+, then use WPF istead.
                b = new Bitmap(1, 1);
                g = Graphics.FromImage(b);
                g.PageUnit = GraphicsUnit.Pixel;
            }
            catch
            {
                return;
            }

            foreach (var cell in _cells)
            {
                if (_worksheet.Column(cell.Start.Column).Hidden)    //Issue 15338
                    continue;

                if (cell.Merge == true || cell.Style.WrapText) continue;
                var fntID = styles.CellXfs[cell.StyleID].FontId;
                Font f;
                if (_fontCache.ContainsKey(fntID))
                {
                    f = _fontCache[fntID];
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

                    _fontCache.Add(fntID, f);
                }
                var ind = styles.CellXfs[cell.StyleID].Indent;
                var textForWidth = cell.TextForWidth;
                var t = textForWidth + (ind > 0 && !string.IsNullOrEmpty(textForWidth) ? new string('_', ind) : "");
                if (t.Length > 32000) t = t.Substring(0, 32000); //Issue
                var size = g.MeasureString(t, f, 10000, StringFormat.GenericDefault);

                double width;
                double r = styles.CellXfs[cell.StyleID].TextRotation;
                if (r <= 0)
                {
                    width = (size.Width + 5) / normalSize;
                }
                else
                {
                    r = (r <= 90 ? r : r - 90);
                    width = (((size.Width - size.Height) * Math.Abs(System.Math.Cos(System.Math.PI * r / 180.0)) + size.Height) + 5) / normalSize;
                }

                foreach (var a in _afAddr)
                {
                    if (a.Collide(cell) != eAddressCollition.No)
                    {
                        width += 2.25;
                        break;
                    }
                }
                _maxWidthMap.AddMeasurement(cell._fromCol, width);
                //if (width > _worksheet.Column(cell._fromCol).Width)
                //{
                //    _worksheet.Column(cell._fromCol).Width = width > MaximumWidth ? MaximumWidth : width;
                //}
            }
        }
    }
}
