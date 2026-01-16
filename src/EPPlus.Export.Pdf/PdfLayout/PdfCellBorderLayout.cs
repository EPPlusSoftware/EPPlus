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
using EPPlus.Graphics;
using EPPlus.Graphics.Math;
using System.Drawing;
using EPPlus.Export.Pdf.PdfSettings;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using EPPlus.Export.Pdf.Pdfhelpers;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCellBorderLayout : Transform , IBorderLayout
    {
        public PdfCellBordersData BorderData;
        public bool IsMerged = false;

        public string range;

        internal PdfTableLayout TableStyle;

        public PdfCellBorderLayout() { }

        public PdfCellBorderLayout(ExcelRangeBase cell, PdfTableLayout tableStyle, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
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
            if (TableStyle != null)
            {
                if ((TableStyle.borderStyleType & TableBorderStyle.Top) != 0)
                {
                    //BorderData.Top = new PdfCellBorderData(LineType.Top);
                    BorderData.Top.BorderStyle = TableStyle.MainStyle.Style.Border.Top.Style == null ? (TableStyle.WholeStyle.Style.Border.Top.Style == null) ? ExcelBorderStyle.None : (ExcelBorderStyle)TableStyle.WholeStyle.Style.Border.Top.Style : (ExcelBorderStyle)TableStyle.MainStyle.Style.Border.Top.Style;
                    BorderData.Top.BorderColor = TableStyle.MainStyle.Style.Border.Top.Style == null ? PdfColor.SetColorFromHex(TableStyle.WholeStyle.Style.Border.Top.Color.LookupColor()) : PdfColor.SetColorFromHex(TableStyle.MainStyle.Style.Border.Top.Color.LookupColor());
                    BorderData.Top.X = LocalPosition.X;
                    BorderData.Top.Y = LocalPosition.Y;
                    BorderData.Top.Width = Size.X;
                    BorderData.Top.Height = Size.Y;
                }
                else if(TableStyle.MainStyle.Style.Border.Horizontal.Style != null)
                {
                    BorderData.Top.BorderStyle = TableStyle.MainStyle.Style.Border.Top.Style == null ? ExcelBorderStyle.None : (ExcelBorderStyle)TableStyle.MainStyle.Style.Border.Top.Style;
                    BorderData.Top.BorderColor = PdfColor.SetColorFromHex(TableStyle.MainStyle.Style.Border.Top.Color.LookupColor());
                    BorderData.Top.X = LocalPosition.X;
                    BorderData.Top.Y = LocalPosition.Y;
                    BorderData.Top.Width = Size.X;
                    BorderData.Top.Height = Size.Y;
                }

                //BorderData.Bottom = new PdfCellBorderData(LineType.Bottom);
                BorderData.Bottom.BorderStyle = TableStyle.MainStyle.Style.Border.Bottom.Style == null ? ExcelBorderStyle.None : (ExcelBorderStyle)TableStyle.MainStyle.Style.Border.Bottom.Style;
                BorderData.Bottom.BorderColor = PdfColor.SetColorFromHex(TableStyle.MainStyle.Style.Border.Bottom.Color.LookupColor());
                BorderData.Bottom.X = LocalPosition.X;
                BorderData.Bottom.Y = LocalPosition.Y;
                BorderData.Bottom.Width = Size.X;
                BorderData.Bottom.Height = Size.Y;
                //BorderData.Left = new PdfCellBorderData(LineType.Left);
                BorderData.Left.BorderStyle = TableStyle.MainStyle.Style.Border.Left.Style == null ? ExcelBorderStyle.None : (ExcelBorderStyle)TableStyle.MainStyle.Style.Border.Left.Style;
                BorderData.Left.BorderColor = PdfColor.SetColorFromHex(TableStyle.MainStyle.Style.Border.Left.Color.LookupColor());
                BorderData.Left.X = LocalPosition.X;
                BorderData.Left.Y = LocalPosition.Y;
                BorderData.Left.Width = Size.X;
                BorderData.Left.Height = Size.Y;
                //BorderData.Right = new PdfCellBorderData(LineType.Right);
                BorderData.Right.BorderStyle = TableStyle.MainStyle.Style.Border.Right.Style == null ? ExcelBorderStyle.None : (ExcelBorderStyle)TableStyle.MainStyle.Style.Border.Right.Style;
                BorderData.Right.BorderColor = PdfColor.SetColorFromHex(TableStyle.MainStyle.Style.Border.Right.Color.LookupColor());
                BorderData.Right.X = LocalPosition.X;
                BorderData.Right.Y = LocalPosition.Y;
                BorderData.Right.Width = Size.X;
                BorderData.Right.Height = Size.Y;
                return;
            }
            if (cell != null)
            {
                //BorderData.Top = new PdfCellBorderData(LineType.Top);
                BorderData.Top.BorderStyle = cell.Style.Border.Top.Style;
                BorderData.Top.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Top.Color.LookupColor(cell.Style.Border));
                BorderData.Top.X = LocalPosition.X;
                BorderData.Top.Y = LocalPosition.Y;
                BorderData.Top.Width = Size.X;
                BorderData.Top.Height = Size.Y;
                //BorderData.Bottom = new PdfCellBorderData(LineType.Bottom);
                BorderData.Bottom.BorderStyle = cell.Style.Border.Bottom.Style;
                BorderData.Bottom.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Bottom.Color.LookupColor(cell.Style.Border));
                BorderData.Bottom.X = LocalPosition.X;
                BorderData.Bottom.Y = LocalPosition.Y;
                BorderData.Bottom.Width = Size.X;
                BorderData.Bottom.Height = Size.Y;
                //BorderData.Left = new PdfCellBorderData(LineType.Left);
                BorderData.Left.BorderStyle = cell.Style.Border.Left.Style;
                BorderData.Left.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Left.Color.LookupColor(cell.Style.Border));
                BorderData.Left.X = LocalPosition.X;
                BorderData.Left.Y = LocalPosition.Y;
                BorderData.Left.Width = Size.X;
                BorderData.Left.Height = Size.Y;
                //BorderData.Right = new PdfCellBorderData(LineType.Right);
                BorderData.Right.BorderStyle = cell.Style.Border.Right.Style;
                BorderData.Right.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Right.Color.LookupColor(cell.Style.Border));
                BorderData.Right.X = LocalPosition.X;
                BorderData.Right.Y = LocalPosition.Y;
                BorderData.Right.Width = Size.X;
                BorderData.Right.Height = Size.Y;
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
    }
}

