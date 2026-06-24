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
using EPPlus.Export.Pdf.Enums;
using EPPlus.Export.Pdf.Helpers;
using EPPlus.Graphics;
using System.Diagnostics;
using System.Drawing;

namespace EPPlus.Export.Pdf.Layout
{
    [DebuggerDisplay("Border: {Name}")]
    internal class PdfCellBorderLayout : Transform
    {
        public PdfCellBordersData BorderData;
        public bool IsMerged = false;
        public MergedCellDrawInfo MergedCellInfo;
        public MergedCellCorners Corners;
        public bool IsPrintTitle;

        public string range;

        internal PdfCellStyle TableStyle;

        public PdfCellBorderLayout() { }

        public PdfCellBorderLayout(PdfCellStyle style, bool isMerged, MergedCellCorners corners, MergedCellDrawInfo info, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y - height, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 3;
            IsMerged = isMerged;
            MergedCellInfo = info;
            Corners = corners;

            BorderData = new PdfCellBordersData();

            BorderData.Top.BorderStyle = style.xfTop.Style == ExcelBorderStyle.None ? ((style.dxfTop != null && style.dxfTop.HasValue) ? (ExcelBorderStyle)style.dxfTop.Style : ExcelBorderStyle.None) : style.xfTop.Style;
            BorderData.Top.BorderColor = style.dxfTop != null ? PdfColor.SetColorFromHex(style.dxfTop.Color.LookupColor(style.dxfTop)) : PdfColor.SetColorFromHex(style.xfTop.Color.LookupColor(style.xfTop));
            BorderData.Top.X = LocalPosition.X;
            BorderData.Top.Y = LocalPosition.Y;
            BorderData.Top.Width = Size.X;
            BorderData.Top.Height = Size.Y;

            BorderData.Bottom.BorderStyle = style.xfBottom.Style == ExcelBorderStyle.None ? ((style.dxfBottom != null && style.dxfBottom.HasValue) ? (ExcelBorderStyle)style.dxfBottom.Style : ExcelBorderStyle.None) : style.xfBottom.Style;
            BorderData.Bottom.BorderColor = style.dxfBottom != null ? PdfColor.SetColorFromHex(style.dxfBottom.Color.LookupColor(style.dxfBottom)) : PdfColor.SetColorFromHex(style.xfBottom.Color.LookupColor(style.xfBottom));
            BorderData.Bottom.X = LocalPosition.X;
            BorderData.Bottom.Y = LocalPosition.Y;
            BorderData.Bottom.Width = Size.X;
            BorderData.Bottom.Height = Size.Y;

            BorderData.Left.BorderStyle = style.xfLeft.Style == ExcelBorderStyle.None ? ((style.dxfLeft != null && style.dxfLeft.HasValue) ? (ExcelBorderStyle)style.dxfLeft.Style : ExcelBorderStyle.None) : style.xfLeft.Style;
            BorderData.Left.BorderColor = style.dxfLeft != null ? PdfColor.SetColorFromHex(style.dxfLeft.Color.LookupColor(style.dxfLeft)) : PdfColor.SetColorFromHex(style.xfLeft.Color.LookupColor(style.xfLeft));
            BorderData.Left.X = LocalPosition.X;
            BorderData.Left.Y = LocalPosition.Y;
            BorderData.Left.Width = Size.X;
            BorderData.Left.Height = Size.Y;

            BorderData.Right.BorderStyle = style.xfRight.Style == ExcelBorderStyle.None ? ((style.dxfRight != null && style.dxfRight.HasValue) ? (ExcelBorderStyle)style.dxfRight.Style : ExcelBorderStyle.None) : style.xfRight.Style;
            BorderData.Right.BorderColor = style.dxfRight != null ? PdfColor.SetColorFromHex(style.dxfRight.Color.LookupColor(style.dxfRight)) : PdfColor.SetColorFromHex(style.xfRight.Color.LookupColor(style.xfRight));
            BorderData.Right.X = LocalPosition.X;
            BorderData.Right.Y = LocalPosition.Y;
            BorderData.Right.Width = Size.X;
            BorderData.Right.Height = Size.Y;

            BorderData.DiagonalUp.BorderStyle = style.DiagonalUp ? style.Diagonal.Style : ExcelBorderStyle.None;
            BorderData.DiagonalUp.BorderColor = style.DiagonalUp ? PdfColor.SetColorFromHex(style.Diagonal.Color.LookupColor(style.Diagonal)) : Color.Transparent;
            BorderData.DiagonalUp.MergedDiagonalWidth = width;
            BorderData.DiagonalUp.MergedDiagonalHeight = height;

            BorderData.DiagonalDown.BorderStyle = style.DiagonalDown ? style.Diagonal.Style : ExcelBorderStyle.None;
            BorderData.DiagonalDown.BorderColor = style.DiagonalDown ? PdfColor.SetColorFromHex(style.Diagonal.Color.LookupColor(style.Diagonal)) : Color.Transparent;
            BorderData.DiagonalDown.MergedDiagonalWidth = width;
            BorderData.DiagonalDown.MergedDiagonalHeight = height;
        }






        //public PdfCellBorderLayout(ExcelRangeBase cell, PdfCellStyle tableStyle, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
        //    : base(x, y-height, width, height, scaleX, scaleY, rotation, parent)
        //{
        //    this.cell = cell;
        //    this.TableStyle = tableStyle;
        //    if (cell != null)
        //    {
        //        IsMerged = cell.Merge;
        //        BorderData = new PdfCellBordersData();
        //    }
        //}

        //public void InitEdgeBorders(ExcelRangeBase cell)
        //{
        //    if (cell != null)
        //    {
        //        //BorderData.Top = new PdfCellBorderData(LineType.Top);
        //        BorderData.Top.BorderStyle = TableStyle.xfTop.Style == ExcelBorderStyle.None ? ((TableStyle.dxfTop != null && TableStyle.dxfTop.HasValue) ? (ExcelBorderStyle)TableStyle.dxfTop.Style : ExcelBorderStyle.None ) : TableStyle.xfTop.Style;
        //        BorderData.Top.BorderColor = TableStyle.dxfTop != null ? PdfColor.SetColorFromHex(TableStyle.dxfTop.Color.LookupColor(TableStyle.dxfTop)) : PdfColor.SetColorFromHex(cell.Style.Border.Top.Color.LookupColor(cell.Style.Border));
        //        BorderData.Top.X = LocalPosition.X;
        //        BorderData.Top.Y = LocalPosition.Y;
        //        BorderData.Top.Width = Size.X;
        //        BorderData.Top.Height = Size.Y;
        //        //BorderData.Bottom = new PdfCellBorderData(LineType.Bottom);
        //        BorderData.Bottom.BorderStyle = TableStyle.xfBottom.Style == ExcelBorderStyle.None ? ((TableStyle.dxfBottom != null && TableStyle.dxfBottom.HasValue) ? (ExcelBorderStyle)TableStyle.dxfBottom.Style : ExcelBorderStyle.None) : TableStyle.xfBottom.Style;
        //        BorderData.Bottom.BorderColor = TableStyle.dxfBottom != null ? PdfColor.SetColorFromHex(TableStyle.dxfBottom.Color.LookupColor(TableStyle.dxfBottom)) : PdfColor.SetColorFromHex(cell.Style.Border.Bottom.Color.LookupColor(cell.Style.Border));
        //        BorderData.Bottom.X = LocalPosition.X;
        //        BorderData.Bottom.Y = LocalPosition.Y;
        //        BorderData.Bottom.Width = Size.X;
        //        BorderData.Bottom.Height = Size.Y;
        //        //BorderData.Left = new PdfCellBorderData(LineType.Left);
        //        BorderData.Left.BorderStyle = TableStyle.xfLeft.Style == ExcelBorderStyle.None ? ((TableStyle.dxfLeft != null && TableStyle.dxfLeft.HasValue) ? (ExcelBorderStyle)TableStyle.dxfLeft.Style : ExcelBorderStyle.None) : TableStyle.xfLeft.Style;
        //        BorderData.Left.BorderColor = TableStyle.dxfLeft != null ? PdfColor.SetColorFromHex(TableStyle.dxfLeft.Color.LookupColor(TableStyle.dxfLeft)) : PdfColor.SetColorFromHex(cell.Style.Border.Left.Color.LookupColor(cell.Style.Border));
        //        BorderData.Left.X = LocalPosition.X;
        //        BorderData.Left.Y = LocalPosition.Y;
        //        BorderData.Left.Width = Size.X;
        //        BorderData.Left.Height = Size.Y;
        //        //BorderData.Right = new PdfCellBorderData(LineType.Right);
        //        BorderData.Right.BorderStyle = TableStyle.xfRight.Style == ExcelBorderStyle.None ? ((TableStyle.dxfRight != null && TableStyle.dxfRight.HasValue) ? (ExcelBorderStyle)TableStyle.dxfRight.Style : ExcelBorderStyle.None) : TableStyle.xfRight.Style;
        //        BorderData.Right.BorderColor = TableStyle.dxfRight != null ? PdfColor.SetColorFromHex(TableStyle.dxfRight.Color.LookupColor(TableStyle.dxfRight)) : PdfColor.SetColorFromHex(cell.Style.Border.Right.Color.LookupColor(cell.Style.Border));
        //        BorderData.Right.X = LocalPosition.X;
        //        BorderData.Right.Y = LocalPosition.Y;
        //        BorderData.Right.Width = Size.X;
        //        BorderData.Right.Height = Size.Y;


        //        ////BorderData.Top = new PdfCellBorderData(LineType.Top);
        //        //BorderData.Top.BorderStyle = cell.Style.Border.Top.Style;
        //        //BorderData.Top.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Top.Color.LookupColor(cell.Style.Border));
        //        //BorderData.Top.X = LocalPosition.X;
        //        //BorderData.Top.Y = LocalPosition.Y;
        //        //BorderData.Top.Width = Size.X;
        //        //BorderData.Top.Height = Size.Y;
        //        ////BorderData.Bottom = new PdfCellBorderData(LineType.Bottom);
        //        //BorderData.Bottom.BorderStyle = cell.Style.Border.Bottom.Style;
        //        //BorderData.Bottom.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Bottom.Color.LookupColor(cell.Style.Border));
        //        //BorderData.Bottom.X = LocalPosition.X;
        //        //BorderData.Bottom.Y = LocalPosition.Y;
        //        //BorderData.Bottom.Width = Size.X;
        //        //BorderData.Bottom.Height = Size.Y;
        //        ////BorderData.Left = new PdfCellBorderData(LineType.Left);
        //        //BorderData.Left.BorderStyle = cell.Style.Border.Left.Style;
        //        //BorderData.Left.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Left.Color.LookupColor(cell.Style.Border));
        //        //BorderData.Left.X = LocalPosition.X;
        //        //BorderData.Left.Y = LocalPosition.Y;
        //        //BorderData.Left.Width = Size.X;
        //        //BorderData.Left.Height = Size.Y;
        //        ////BorderData.Right = new PdfCellBorderData(LineType.Right);
        //        //BorderData.Right.BorderStyle = cell.Style.Border.Right.Style;
        //        //BorderData.Right.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Right.Color.LookupColor(cell.Style.Border));
        //        //BorderData.Right.X = LocalPosition.X;
        //        //BorderData.Right.Y = LocalPosition.Y;
        //        //BorderData.Right.Width = Size.X;
        //        //BorderData.Right.Height = Size.Y;
        //    }
        //}

        //public void InitDiagonalBorders(ExcelRangeBase cell, double width, double height)
        //{
        //    if (cell != null)
        //    {
        //        //BorderData.DiagonalUp = new PdfCellBorderData(LineType.DiagonalUp);
        //        BorderData.DiagonalUp.BorderStyle = cell.Style.Border.DiagonalUp ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
        //        BorderData.DiagonalUp.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
        //        BorderData.DiagonalUp.MergedDiagonalWidth = width;
        //        BorderData.DiagonalUp.MergedDiagonalHeight = height;
        //        //BorderData.DiagonalDown = new PdfCellBorderData(LineType.DiagonalDown);
        //        BorderData.DiagonalDown.BorderStyle = cell.Style.Border.DiagonalDown ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
        //        BorderData.DiagonalDown.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
        //        BorderData.DiagonalDown.MergedDiagonalWidth = width;
        //        BorderData.DiagonalDown.MergedDiagonalHeight = height;
        //    }
        //}

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



        /*
         * Following methods needs to be refactored into 1 method.
         * Less code is more good.
         */
    }
}

