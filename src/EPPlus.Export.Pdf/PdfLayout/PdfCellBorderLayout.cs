/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfSettings;
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using OfficeOpenXml;
using OfficeOpenXml.Core.Worksheet.XmlWriter;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.Table;
using OfficeOpenXml.Table;
using System.Drawing;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCellBorderLayout : Transform , IBorderLayout
    {
        public PdfCellBordersData BorderData;
        public bool IsMerged = false;

        public string range;

        internal PdfCellStyleOverride TableStyle;

        public PdfCellBorderLayout() { }

        public PdfCellBorderLayout(ExcelRangeBase cell, PdfCellStyleOverride tableStyle, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y-height, width, height, scaleX, scaleY, rotation, parent)
        {
            this.cell = cell;
            this.TableStyle = tableStyle;
            if (cell != null)
            {
                IsMerged = cell.Merge;
                BorderData = new PdfCellBordersData();
            }
        }

        public void InitEdgeBorders(ExcelRangeBase cell)
        {
            if (cell != null)
            {
                //BorderData.Top = new PdfCellBorderData(LineType.Top);
                BorderData.Top.BorderStyle = TableStyle.xfTop.Style == ExcelBorderStyle.None ? ((TableStyle.dxfTop != null && TableStyle.dxfTop.HasValue) ? (ExcelBorderStyle)TableStyle.dxfTop.Style : ExcelBorderStyle.None ) : TableStyle.xfTop.Style;
                BorderData.Top.BorderColor = TableStyle.dxfTop != null ? PdfColor.SetColorFromHex(TableStyle.dxfTop.Color.LookupColor(TableStyle.dxfTop)) : PdfColor.SetColorFromHex(cell.Style.Border.Top.Color.LookupColor(cell.Style.Border));
                BorderData.Top.X = LocalPosition.X;
                BorderData.Top.Y = LocalPosition.Y;
                BorderData.Top.Width = Size.X;
                BorderData.Top.Height = Size.Y;
                //BorderData.Bottom = new PdfCellBorderData(LineType.Bottom);
                BorderData.Bottom.BorderStyle = TableStyle.xfBottom.Style == ExcelBorderStyle.None ? ((TableStyle.dxfBottom != null && TableStyle.dxfBottom.HasValue) ? (ExcelBorderStyle)TableStyle.dxfBottom.Style : ExcelBorderStyle.None) : TableStyle.xfBottom.Style;
                BorderData.Bottom.BorderColor = TableStyle.dxfBottom != null ? PdfColor.SetColorFromHex(TableStyle.dxfBottom.Color.LookupColor(TableStyle.dxfBottom)) : PdfColor.SetColorFromHex(cell.Style.Border.Bottom.Color.LookupColor(cell.Style.Border));
                BorderData.Bottom.X = LocalPosition.X;
                BorderData.Bottom.Y = LocalPosition.Y;
                BorderData.Bottom.Width = Size.X;
                BorderData.Bottom.Height = Size.Y;
                //BorderData.Left = new PdfCellBorderData(LineType.Left);
                BorderData.Left.BorderStyle = TableStyle.xfLeft.Style == ExcelBorderStyle.None ? ((TableStyle.dxfLeft != null && TableStyle.dxfLeft.HasValue) ? (ExcelBorderStyle)TableStyle.dxfLeft.Style : ExcelBorderStyle.None) : TableStyle.xfLeft.Style;
                BorderData.Left.BorderColor = TableStyle.dxfLeft != null ? PdfColor.SetColorFromHex(TableStyle.dxfLeft.Color.LookupColor(TableStyle.dxfLeft)) : PdfColor.SetColorFromHex(cell.Style.Border.Left.Color.LookupColor(cell.Style.Border));
                BorderData.Left.X = LocalPosition.X;
                BorderData.Left.Y = LocalPosition.Y;
                BorderData.Left.Width = Size.X;
                BorderData.Left.Height = Size.Y;
                //BorderData.Right = new PdfCellBorderData(LineType.Right);
                BorderData.Right.BorderStyle = TableStyle.xfRight.Style == ExcelBorderStyle.None ? ((TableStyle.dxfRight != null && TableStyle.dxfRight.HasValue) ? (ExcelBorderStyle)TableStyle.dxfRight.Style : ExcelBorderStyle.None) : TableStyle.xfRight.Style;
                BorderData.Right.BorderColor = TableStyle.dxfRight != null ? PdfColor.SetColorFromHex(TableStyle.dxfRight.Color.LookupColor(TableStyle.dxfRight)) : PdfColor.SetColorFromHex(cell.Style.Border.Right.Color.LookupColor(cell.Style.Border));
                BorderData.Right.X = LocalPosition.X;
                BorderData.Right.Y = LocalPosition.Y;
                BorderData.Right.Width = Size.X;
                BorderData.Right.Height = Size.Y;


                ////BorderData.Top = new PdfCellBorderData(LineType.Top);
                //BorderData.Top.BorderStyle = cell.Style.Border.Top.Style;
                //BorderData.Top.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Top.Color.LookupColor(cell.Style.Border));
                //BorderData.Top.X = LocalPosition.X;
                //BorderData.Top.Y = LocalPosition.Y;
                //BorderData.Top.Width = Size.X;
                //BorderData.Top.Height = Size.Y;
                ////BorderData.Bottom = new PdfCellBorderData(LineType.Bottom);
                //BorderData.Bottom.BorderStyle = cell.Style.Border.Bottom.Style;
                //BorderData.Bottom.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Bottom.Color.LookupColor(cell.Style.Border));
                //BorderData.Bottom.X = LocalPosition.X;
                //BorderData.Bottom.Y = LocalPosition.Y;
                //BorderData.Bottom.Width = Size.X;
                //BorderData.Bottom.Height = Size.Y;
                ////BorderData.Left = new PdfCellBorderData(LineType.Left);
                //BorderData.Left.BorderStyle = cell.Style.Border.Left.Style;
                //BorderData.Left.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Left.Color.LookupColor(cell.Style.Border));
                //BorderData.Left.X = LocalPosition.X;
                //BorderData.Left.Y = LocalPosition.Y;
                //BorderData.Left.Width = Size.X;
                //BorderData.Left.Height = Size.Y;
                ////BorderData.Right = new PdfCellBorderData(LineType.Right);
                //BorderData.Right.BorderStyle = cell.Style.Border.Right.Style;
                //BorderData.Right.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Right.Color.LookupColor(cell.Style.Border));
                //BorderData.Right.X = LocalPosition.X;
                //BorderData.Right.Y = LocalPosition.Y;
                //BorderData.Right.Width = Size.X;
                //BorderData.Right.Height = Size.Y;
            }
        }

        public void InitDiagonalBorders(ExcelRangeBase cell, double width, double height)
        {
            if (cell != null)
            {
                //BorderData.DiagonalUp = new PdfCellBorderData(LineType.DiagonalUp);
                BorderData.DiagonalUp.BorderStyle = cell.Style.Border.DiagonalUp ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalUp.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
                BorderData.DiagonalUp.MergedDiagonalWidth = width;
                BorderData.DiagonalUp.MergedDiagonalHeight = height;
                //BorderData.DiagonalDown = new PdfCellBorderData(LineType.DiagonalDown);
                BorderData.DiagonalDown.BorderStyle = cell.Style.Border.DiagonalDown ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalDown.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
                BorderData.DiagonalDown.MergedDiagonalWidth = width;
                BorderData.DiagonalDown.MergedDiagonalHeight = height;
            }
        }

        public void UpdateLocalBorderPosition()
        {
            //if (cell.Address == FirstCellInMerge)
            //{
            //    BorderData.Top.X = LocalPosition.X;
            //    BorderData.Top.Y = LocalPosition.Y + Size.Y - BorderData.Top.Height;
            //}
            //else
            //{
            //    BorderData.Top.X = LocalPosition.X;
            //    BorderData.Top.Y = LocalPosition.Y;
            //}
            //BorderData.Bottom.X = LocalPosition.X;
            //BorderData.Bottom.Y = LocalPosition.Y;
            //if (cell.Address == FirstCellInMerge)
            //{
            //    BorderData.Left.X = LocalPosition.X;
            //    BorderData.Left.Y = LocalPosition.Y + Size.Y - BorderData.Left.Height;
            //}
            //else
            //{
            //    BorderData.Left.X = LocalPosition.X;
            //    BorderData.Left.Y = LocalPosition.Y;
            //}
            //BorderData.Right.X = LocalPosition.X;
            //BorderData.Right.Y = LocalPosition.Y;
        }


        //Get Methods fort border styles
        /*
             Whole Table
             First Column Stripe
             Second Column Stripe
             First Row Stripe
             Second Row Stripe
             Last Column
             First Column
             Header Row
             Total Row
             First Header Cell
             Last Header Cell
             First Total Cell
             Last Total Cell
        */

        public static ExcelDxfBorderItem GetTopBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var top = tableRow == 0 ? tableStyle.WholeTable.Style.Border.Top : tableStyle.WholeTable.Style.Border.Horizontal;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Top.HasValue)
                {
                    top = tableStyle.HeaderRow.Style.Border.Top;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Top.HasValue)
                {
                    top = tableStyle.FirstHeaderCell.Style.Border.Top;
                }
                if (tableCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Top.HasValue)
                {
                    top = tableStyle.LastHeaderCell.Style.Border.Top;
                }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Top.HasValue)
                {
                    top = tableStyle.TotalRow.Style.Border.Top;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Top.HasValue)
                {
                    top = tableStyle.FirstTotalCell.Style.Border.Top;
                }
                if (tableCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Top.HasValue)
                {
                    top = tableStyle.LastTotalCell.Style.Border.Top;
                }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Top.HasValue)
                {
                    top = tableStyle.FirstColumnStripe.Style.Border.Top;
                }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Top.HasValue)
                {
                    top = tableStyle.SecondColumnStripe.Style.Border.Top;
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Top.HasValue)
                {
                    top = tableStyle.FirstRowStripe.Style.Border.Top;
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Top.HasValue)
                {
                    top = tableStyle.SecondRowStripe.Style.Border.Top;
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Top.HasValue)
                {
                    top = tableStyle.LastColumn.Style.Border.Top;
                }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Top.HasValue)
                {
                    top = tableStyle.FirstColumn.Style.Border.Top;
                }
                //we are inside so we use horizontal as fallback from whole table else we use top.
                //if (tableStyle.WholeTable.Style.Border.Horizontal.HasValue)
                //{
                //    top = tableStyle.WholeTable.Style.Border.Horizontal;
                //}
            }
            return top;
        }
        public static ExcelDxfBorderItem GetBottomBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var bottom = range._toRow == cell._fromRow ? tableStyle.WholeTable.Style.Border.Bottom : null;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.HeaderRow.Style.Border.Bottom;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.FirstHeaderCell.Style.Border.Bottom;
                }
                if (tableCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.LastHeaderCell.Style.Border.Bottom;
                }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.TotalRow.Style.Border.Bottom;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.FirstTotalCell.Style.Border.Bottom;
                }
                if (tableCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.LastTotalCell.Style.Border.Bottom;
                }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.FirstColumnStripe.Style.Border.Bottom;
                }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.SecondColumnStripe.Style.Border.Bottom;
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Bottom.HasValue && (tableRow & 1) != 0)
                {
                    //check if cell is first row stipe eligble
                    bottom = tableStyle.FirstRowStripe.Style.Border.Bottom;
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Bottom.HasValue && (tableRow & 1) == 0)
                {
                    //check if cell is second row stipe eligble
                    bottom = tableStyle.SecondRowStripe.Style.Border.Bottom;
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.LastColumn.Style.Border.Bottom;
                }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Bottom.HasValue)
                {
                    bottom = tableStyle.FirstColumn.Style.Border.Bottom;
                }
                ////we are inside so we use horizontal as fallback from whole table else we use Bottom.
                //if (tableStyle.WholeTable.Style.Border.Horizontal.HasValue)
                //{
                //    bottom = tableStyle.WholeTable.Style.Border.Horizontal;
                //}
            }
            return bottom;
        }
        public static ExcelDxfBorderItem GetLeftBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var left = tableCol == 0 ? tableStyle.WholeTable.Style.Border.Left : tableStyle.WholeTable.Style.Border.Vertical;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Left.HasValue)
                {
                    left = tableStyle.HeaderRow.Style.Border.Left;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Left.HasValue)
                {
                    left = tableStyle.FirstHeaderCell.Style.Border.Left;
                }
                if (tableCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Left.HasValue)
                {
                    left = tableStyle.LastHeaderCell.Style.Border.Left;
                }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Left.HasValue)
                {
                    left = tableStyle.TotalRow.Style.Border.Left;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Left.HasValue)
                {
                    left = tableStyle.FirstTotalCell.Style.Border.Left;
                }
                if (tableCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Left.HasValue)
                {
                    left = tableStyle.LastTotalCell.Style.Border.Left;
                }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Left.HasValue)
                {
                    left = tableStyle.FirstColumnStripe.Style.Border.Left;
                }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Left.HasValue)
                {
                    left = tableStyle.SecondColumnStripe.Style.Border.Left;
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Left.HasValue)
                {
                    left = tableStyle.FirstRowStripe.Style.Border.Left;
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Left.HasValue)
                {
                    left = tableStyle.SecondRowStripe.Style.Border.Left;
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Left.HasValue)
                {
                    left = tableStyle.LastColumn.Style.Border.Left;
                }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Left.HasValue)
                {
                    left = tableStyle.FirstColumn.Style.Border.Left;
                }
                //we are inside so we use Vertical as fallback from whole table else we use Left.
                if (tableStyle.WholeTable.Style.Border.Vertical.HasValue)
                {
                    left = tableStyle.WholeTable.Style.Border.Vertical;
                }
            }
            return left;
        }
        public static ExcelDxfBorderItem GetRightBorderItem(ExcelRangeBase cell, ExcelBorderItem xfBorder, ExcelTable table, ExcelTableNamedStyle tableStyle)
        {
            var range = table.Range;
            int tableRow = cell._fromRow - range._fromRow;
            int tableCol = cell._fromCol - range._fromCol;
            var right = cell._fromCol == range._toCol ? tableStyle.WholeTable.Style.Border.Right : tableStyle.WholeTable.Style.Border.Vertical;
            if (table.ShowHeader && tableRow == 0)
            {
                if (tableStyle.HeaderRow.Style.Border.Right.HasValue)
                {
                    right = tableStyle.HeaderRow.Style.Border.Right;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstHeaderCell.Style.Border.Right.HasValue)
                {
                    right = tableStyle.FirstHeaderCell.Style.Border.Right;
                }
                if (tableCol == range._toCol && table.ShowLastColumn && tableStyle.LastHeaderCell.Style.Border.Right.HasValue)
                {
                    right = tableStyle.LastHeaderCell.Style.Border.Right;
                }
            }
            else if (table.ShowTotal && cell._fromRow == range._toRow)
            {
                if (tableStyle.TotalRow.Style.Border.Right.HasValue)
                {
                    right = tableStyle.TotalRow.Style.Border.Right;
                }
                if (tableCol == 0 && table.ShowFirstColumn && tableStyle.FirstTotalCell.Style.Border.Right.HasValue)
                {
                    right = tableStyle.FirstTotalCell.Style.Border.Right;
                }
                if (cell._fromCol == range._toCol && table.ShowLastColumn && tableStyle.LastTotalCell.Style.Border.Right.HasValue)
                {
                    right = tableStyle.LastTotalCell.Style.Border.Right;
                }
            }
            else
            {
                if (table.ShowColumnStripes && tableStyle.FirstColumnStripe.Style.Border.Right.HasValue)
                {
                    right = tableStyle.FirstColumnStripe.Style.Border.Right;
                }
                if (table.ShowColumnStripes && tableStyle.SecondColumnStripe.Style.Border.Right.HasValue)
                {
                    right = tableStyle.SecondColumnStripe.Style.Border.Right;
                }
                if (table.ShowRowStripes && tableStyle.FirstRowStripe.Style.Border.Right.HasValue)
                {
                    right = tableStyle.FirstRowStripe.Style.Border.Right;
                }
                if (table.ShowRowStripes && tableStyle.SecondRowStripe.Style.Border.Right.HasValue)
                {
                    right = tableStyle.SecondRowStripe.Style.Border.Right;
                }
                if (table.ShowLastColumn && tableStyle.LastColumn.Style.Border.Right.HasValue)
                {
                    right = tableStyle.LastColumn.Style.Border.Right;
                }
                if (table.ShowFirstColumn && tableStyle.FirstColumn.Style.Border.Right.HasValue)
                {
                    right = tableStyle.FirstColumn.Style.Border.Right;
                }
                //we are inside so we use Vertical as fallback from whole table else we use Right.
                if (tableStyle.WholeTable.Style.Border.Vertical.HasValue)
                {
                    right = tableStyle.WholeTable.Style.Border.Vertical;
                }
            }
            return right;
        }
    }
}

