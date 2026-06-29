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

        public PdfCellBorderLayout(bool isMerged, MergedCellCorners corners, MergedCellDrawInfo info, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y - height, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 3;
            IsMerged = isMerged;
            MergedCellInfo = info;
            Corners = corners;
            BorderData = new PdfCellBordersData();

            BorderData.Top.X = LocalPosition.X;
            BorderData.Top.Y = LocalPosition.Y;
            BorderData.Top.Width = Size.X;
            BorderData.Top.Height = Size.Y;

            BorderData.Bottom.X = LocalPosition.X;
            BorderData.Bottom.Y = LocalPosition.Y;
            BorderData.Bottom.Width = Size.X;
            BorderData.Bottom.Height = Size.Y;

            BorderData.Left.X = LocalPosition.X;
            BorderData.Left.Y = LocalPosition.Y;
            BorderData.Left.Width = Size.X;
            BorderData.Left.Height = Size.Y;

            BorderData.Right.X = LocalPosition.X;
            BorderData.Right.Y = LocalPosition.Y;
            BorderData.Right.Width = Size.X;
            BorderData.Right.Height = Size.Y;

            BorderData.DiagonalUp.MergedDiagonalWidth = width;
            BorderData.DiagonalUp.MergedDiagonalHeight = height;

            BorderData.DiagonalDown.MergedDiagonalWidth = width;
            BorderData.DiagonalDown.MergedDiagonalHeight = height;
        }

        internal void SetStyle(ExcelBorderStyle topBorderStyle, Color topBorderColor,
                               ExcelBorderStyle bottomBorderStyle, Color bottomBorderColor,
                               ExcelBorderStyle leftBorderStyle, Color leftBorderColor,
                               ExcelBorderStyle rightBorderStyle, Color rightBorderColor,
                               ExcelBorderStyle diagUpBorderStyle, Color diagUpBorderColor,
                               ExcelBorderStyle diagdDownBorderStyle, Color diagDownBorderColor)
        {
            BorderData.Top.BorderStyle = topBorderStyle;
            BorderData.Top.BorderColor = topBorderColor;

            BorderData.Bottom.BorderStyle = bottomBorderStyle;
            BorderData.Bottom.BorderColor = bottomBorderColor;

            BorderData.Left.BorderStyle = leftBorderStyle;
            BorderData.Left.BorderColor = leftBorderColor;

            BorderData.Right.BorderStyle = rightBorderStyle;
            BorderData.Right.BorderColor = rightBorderColor;

            BorderData.DiagonalUp.BorderStyle = diagUpBorderStyle;
            BorderData.DiagonalUp.BorderColor = diagUpBorderColor;

            BorderData.DiagonalDown.BorderStyle = diagdDownBorderStyle;
            BorderData.DiagonalDown.BorderColor = diagDownBorderColor;
        }
    }
}

