using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml;

namespace OfficeOpenXml.Drawing.Vml
{
    public class BaseSizing : XmlHelper
    {
        internal const float STANDARD_DPI = 96;
        /// <summary>
        /// The ratio between EMU and Pixels
        /// </summary>
        public const int EMU_PER_PIXEL = 9525;
        /// <summary>
        /// The ratio between EMU and Points
        /// </summary>
        public const int EMU_PER_POINT = 12700;
        /// <summary>
        /// The ratio between EMU and centimeters
        /// </summary>
        public const int EMU_PER_CM = 360000;
        /// <summary>
        /// The ratio between EMU and milimeters
        /// </summary>
        public const int EMU_PER_MM = 3600000;
        /// <summary>
        /// The ratio between EMU and US Inches
        /// </summary>
        public const int EMU_PER_US_INCH = 914400;
        /// <summary>
        /// The ratio between EMU and pica
        /// </summary>
        public const int EMU_PER_PICA = EMU_PER_US_INCH / 6;

        /// <summary>
        /// Top Left position, if the shape is of the absolute anchor type
        /// </summary>
        public ExcelDrawingCoordinate Position
        {
            get;
            private set;
        }

        internal double _width = double.MinValue, _height = double.MinValue, _top = double.MinValue, _left = double.MinValue;

        ExcelWorksheet _ws;

        internal bool _doNotAdjust = false;

        internal BaseSizing(XmlNode topNode, XmlNamespaceManager ns, ExcelWorksheet ws) : base(ns, topNode)
        {
            _ws = ws;
        }

        public eEditAs CellAnchor
        {
            get;
            protected set;
        }

        internal ExcelPositionBase From
        {
            get;
            set;
        }

        internal ExcelPositionBase To
        {
            get;
            set;
        }

        /// <summary>
        /// The extent of the shape, if the shape is of the one- or absolute- anchor type.
        /// Otherwise this propery is set to null
        /// </summary>
        public ExcelDrawingSize Size
        {
            get;
            private set;
        }


        internal void GetFromBounds(out int fromRow, out int fromRowOff, out int fromCol, out int fromColOff)
        {
            if (CellAnchor == eEditAs.Absolute)
            {
                GetToRowFromPixels(Position.Y, out fromRow, out fromRowOff);
                GetToColumnFromPixels(Position.X, out fromCol, out fromColOff);
            }
            else
            {
                fromRow = From.Row;
                fromRowOff = From.RowOff;
                fromCol = From.Column;
                fromColOff = From.ColumnOff;
            }
        }
        internal void GetToBounds(out int toRow, out int toRowOff, out int toCol, out int toColOff)
        {
            if (CellAnchor == eEditAs.Absolute)
            {
                GetToRowFromPixels((Position.Y + Size.Height) / EMU_PER_PIXEL, out toRow, out toRowOff);
                GetToColumnFromPixels(Position.X + Size.Width / EMU_PER_PIXEL, out toCol, out toColOff);
            }
            else
            {
                if (CellAnchor == eEditAs.TwoCell)
                {
                    toRow = To.Row;
                    toRowOff = To.RowOff;
                    toCol = To.Column;
                    toColOff = To.ColumnOff;
                }
                else
                {
                    GetToRowFromPixels(Size.Height / EMU_PER_PIXEL, out toRow, out toRowOff, From.Row, From.RowOff);
                    GetToColumnFromPixels(Size.Width / EMU_PER_PIXEL, out toCol, out toColOff, From.Column, From.ColumnOff);
                }
            }
        }
        internal int GetPixelLeft()
        {
            int pix;
            if (CellAnchor == eEditAs.Absolute)
            {
                pix = Position.X / EMU_PER_PIXEL;
            }
            else
            {
                ExcelWorksheet ws = _ws;
                decimal mdw = ws.Workbook.MaxFontWidth;

                pix = 0;
                for (int col = 0; col < From.Column; col++)
                {
                    pix += ws.GetColumnWidthPixels(col, mdw);
                }
                pix += From.ColumnOff / EMU_PER_PIXEL;
            }

            return pix;
        }
        internal int GetPixelTop()
        {
            int pix;
            if (CellAnchor == eEditAs.Absolute)
            {
                pix = Position.Y / EMU_PER_PIXEL;
            }
            else
            {
                pix = 0;
                var cache = _ws.RowHeightCache;
                for (int row = 0; row < From.Row; row++)
                {
                    lock (cache)
                    {
                        if (!cache.ContainsKey(row))
                        {
                            cache.Add(row, _ws.GetRowHeight(row + 1));
                        }
                    }
                    pix += (int)(cache[row] / 0.75);
                }
                pix += From.RowOff / EMU_PER_PIXEL;
            }
            return pix;
        }

        internal double GetPixelWidth()
        {
            double pix;
            if (CellAnchor == eEditAs.TwoCell)
            {
                ExcelWorksheet ws = _ws;
                decimal mdw = ws.Workbook.MaxFontWidth;

                pix = -From.ColumnOff / (double)EMU_PER_PIXEL;
                for (int col = From.Column + 1; col <= To.Column; col++)
                {
                    pix += (double)decimal.Truncate(((256 * ws.GetColumnWidth(col) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                }

                var w = (double)decimal.Truncate(((256 * ws.GetColumnWidth(To.Column + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
                pix += Math.Min(w, Convert.ToDouble(To.ColumnOff) / EMU_PER_PIXEL);
            }
            else
            {
                pix = Size.Width / (double)EMU_PER_PIXEL;
            }
            return pix;
        }
        internal double GetPixelHeight()
        {
            double pix;
            if (CellAnchor == eEditAs.TwoCell)
            {
                ExcelWorksheet ws = _ws;

                pix = -(From.RowOff / (double)EMU_PER_PIXEL);
                for (int row = From.Row + 1; row <= To.Row; row++)
                {
                    pix += ws.GetRowHeight(row) / 0.75;
                }
                var h = ws.GetRowHeight(To.Row + 1) / 0.75;
                pix += Math.Min(h, Convert.ToDouble(To.RowOff) / EMU_PER_PIXEL);
            }
            else
            {
                pix = Size.Height / (double)EMU_PER_PIXEL;
            }
            return pix;
        }

        //internal void SetPixelTop(double pixels)
        //{
        //    _doNotAdjust = true;
        //    if (CellAnchor == eEditAs.Absolute)
        //    {
        //        Position.Y = (int)(pixels * EMU_PER_PIXEL);
        //    }
        //    else
        //    {
        //        CalcRowFromPixelTop(pixels, out int row, out int rowOff);
        //        From.Row = row;
        //        From.RowOff = rowOff;
        //    }
        //    _top = pixels;
        //    _doNotAdjust = false;
        //}

        //internal void CalcRowFromPixelTop(double pixels, out int row, out int rowOff)
        //{
        //    ExcelWorksheet ws = _drawings.Worksheet;
        //    decimal mdw = ws.Workbook.MaxFontWidth;
        //    double prevPix = 0;
        //    double pix = ws.GetRowHeight(1) / 0.75;
        //    int r = 2;
        //    while (pix < pixels)
        //    {
        //        prevPix = pix;
        //        pix += (int)(ws.GetRowHeight(r++) / 0.75);
        //    }

        //    if (pix == pixels)
        //    {
        //        row = r - 1;
        //        rowOff = 0;
        //    }
        //    else
        //    {
        //        row = r - 2;
        //        rowOff = (int)(pixels - prevPix) * EMU_PER_PIXEL;
        //    }
        //}

        internal void SetPixelLeft(double pixels)
        {
            _doNotAdjust = true;
            if (CellAnchor == eEditAs.Absolute)
            {
                Position.X = (int)(pixels * EMU_PER_PIXEL);
            }
            else
            {
                CalcColFromPixelLeft(pixels, out int col, out int colOff);
                From.Column = col;
                From.ColumnOff = colOff;
            }
            _doNotAdjust = false;

            _left = pixels;
        }
        internal void CalcColFromPixelLeft(double pixels, out int column, out int columnOff)
        {

            ExcelWorksheet ws = _ws;
            decimal mdw = ws.Workbook.MaxFontWidth;
            double prevPix = 0;
            double pix = (int)decimal.Truncate(((256 * ws.GetColumnWidth(1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            int col = 2;

            while (pix < pixels)
            {
                prevPix = pix;
                pix += (int)decimal.Truncate(((256 * ws.GetColumnWidth(col++) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            }
            if (pix == pixels)
            {
                column = col - 1;
                columnOff = 0;
            }
            else
            {
                column = col - 2;
                columnOff = (int)(pixels - prevPix) * EMU_PER_PIXEL;
            }
        }
        internal void SetPixelHeight(double pixels)
        {
            if (CellAnchor == eEditAs.TwoCell)
            {
                _doNotAdjust = true;
                GetToRowFromPixels(pixels, out int toRow, out int pixOff);
                To.Row = toRow;
                To.RowOff = pixOff;
                _doNotAdjust = false;
            }
            else
            {
                Size.Height = (long)Math.Round(pixels * EMU_PER_PIXEL);
            }
        }

        internal void GetToRowFromPixels(double pixels, out int toRow, out int rowOff, int fromRow = -1, int fromRowOff = -1)
        {
            if (fromRow < 0)
            {
                fromRow = From.Row;
                fromRowOff = From.RowOff;
            }
            ExcelWorksheet ws = _ws;
            var pixOff = pixels - ((ws.GetRowHeight(fromRow + 1) / 0.75) - (fromRowOff / (double)EMU_PER_PIXEL));
            double prevPixOff = pixels;
            int row = fromRow + 1;

            while (pixOff >= 0)
            {
                prevPixOff = pixOff;
                pixOff -= (ws.GetRowHeight(++row) / 0.75);
            }
            toRow = row - 1;
            if (fromRow == toRow)
            {
                rowOff = (int)(fromRowOff + (pixels) * EMU_PER_PIXEL);
            }
            else
            {
                rowOff = (int)(prevPixOff * EMU_PER_PIXEL);
            }
        }

        internal void SetPixelWidth(double pixels)
        {
            if (CellAnchor == eEditAs.TwoCell)
            {
                _doNotAdjust = true;
                GetToColumnFromPixels(pixels, out int col, out int pixOff);

                To.Column = col - 2;
                To.ColumnOff = pixOff * EMU_PER_PIXEL;
                _doNotAdjust = false;
            }
            else
            {
                Size.Width = (int)Math.Round(pixels * EMU_PER_PIXEL);
            }
        }

        internal void GetToColumnFromPixels(double pixels, out int col, out int colOff, int fromColumn = -1, int fromColumnOff = -1)
        {
            ExcelWorksheet ws = _ws;
            decimal mdw = ws.Workbook.MaxFontWidth;
            if (fromColumn < 0)
            {
                fromColumn = From.Column;
                fromColumnOff = From.ColumnOff;
            }
            double pixOff = pixels - (double)(decimal.Truncate(((256 * ws.GetColumnWidth(fromColumn + 1) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw) - fromColumnOff / EMU_PER_PIXEL);
            double offset = (double)fromColumnOff / EMU_PER_PIXEL + pixels;
            col = fromColumn + 2;
            while (pixOff >= 0)
            {
                offset = pixOff;
                pixOff -= (double)decimal.Truncate(((256 * ws.GetColumnWidth(col++) + decimal.Truncate(128 / (decimal)mdw)) / 256) * mdw);
            }
            colOff = (int)offset;
        }
    }
}
