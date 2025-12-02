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
    internal class PdfCellBorderLayout : Transform, ILayout
    {
        public PdfCellBordersData BorderData;

        public PdfCellBorderLayout() { }

        public PdfCellBorderLayout(ExcelRangeBase cell, ExcelRangeBase dimension, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y, width, height, scaleX, scaleY, rotation, parent)
        {
            if (cell != null)
            {
                BorderData = new PdfCellBordersData();
                BorderData.Top.BorderStyle = cell.Style.Border.Top.Style;
                BorderData.Top.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Top.Color.LookupColor(cell.Style.Border));
                BorderData.Bottom.BorderStyle = cell.Style.Border.Bottom.Style;
                BorderData.Bottom.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Bottom.Color.LookupColor(cell.Style.Border));
                BorderData.Left.BorderStyle = cell.Style.Border.Left.Style;
                BorderData.Left.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Left.Color.LookupColor(cell.Style.Border));
                BorderData.Right.BorderStyle = cell.Style.Border.Right.Style;
                BorderData.Right.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Right.Color.LookupColor(cell.Style.Border));
                BorderData.DiagonalUp.BorderStyle = cell.Style.Border.DiagonalUp ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalUp.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
                BorderData.DiagonalDown.BorderStyle = cell.Style.Border.DiagonalDown ? cell.Style.Border.Diagonal.Style : ExcelBorderStyle.None;
                BorderData.DiagonalDown.BorderColor = PdfColor.SetColorFromHex(cell.Style.Border.Diagonal.Color.LookupColor(cell.Style.Border));
            }
        }

        public void ConvertCoordinates(PdfPageSettings pageSettings)
        {
            LocalPosition = new Vector2(LocalPosition.X, pageSettings.PageSize.HeightPu - System.Math.Abs(LocalPosition.Y) - Size.Y);
        }
    }
}

