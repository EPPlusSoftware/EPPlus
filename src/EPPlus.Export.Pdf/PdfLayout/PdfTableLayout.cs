using System;
using EPPlus.Export.Pdf.Pdfhelpers;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal enum TableCellStyle
    {
        Header,
        OddRow,
        EvenRow,
        TotalRow,
        FirstColumn,
        LastColumn,
        OddColumn,
        EvenColumn,
        WholeTable
    }

    [Flags]
    public enum TableBorderStyle
    {
        None = 0,
        Top = 1 << 0,
        Bottom = 1 << 1,
        Left = 1 << 2,
        Right = 1 << 3,
        Horizontal = 1 << 4,
        Vertical = 1 << 5
    }

    internal class PdfCellStyleOverride
    {
        internal TableCellStyle TableCellStyleType { get; set; }
        internal ExcelTableStyleElement MainStyleFill;
        internal ExcelTableStyleElement WholeStyleFill;

        internal ExcelFill xfFill;
        internal ExcelDxfFill dxfFill;


        internal ExcelBorderItem Top;
        internal ExcelBorderItem Bottom;
        internal ExcelBorderItem Left;
        internal ExcelBorderItem Right;
        internal ExcelBorderItem DiagonalUp;
        internal ExcelBorderItem DiagonalDown;
        internal ExcelFont Font;


        internal TableBorderStyle borderStyleType = TableBorderStyle.None;



        internal ExcelFill GetAppliedFill()
        {
            ExcelFill Fill = new ExcelFill(xfFill._styles, xfFill._styles.PropertyChange, xfFill._positionID, xfFill._address, -1);
            Fill.PatternType = xfFill.PatternType;
            if (Fill.PatternType != ExcelFillStyle.None)
            {
                Fill.BackgroundColor.SetColor(PdfColor.SetColorFromHex(xfFill.BackgroundColor.LookupColor()));
                Fill.PatternColor.SetColor(PdfColor.SetColorFromHex(xfFill.PatternColor.LookupColor()));
            }
            if (xfFill.HasGradient)
            {
                Fill.PatternType = ExcelFillStyle.None;
                var g = xfFill.Gradient;
                Fill.Gradient.Type = xfFill.Gradient.Type;
                Fill.Gradient.Color1.SetColor(PdfColor.SetColorFromHex(xfFill.Gradient.Color1.LookupColor()));
                Fill.Gradient.Color2.SetColor(PdfColor.SetColorFromHex(xfFill.Gradient.Color2.LookupColor()));
                Fill.Gradient.Color3.SetColor(PdfColor.SetColorFromHex(xfFill.Gradient.Color3.LookupColor()));
                Fill.Gradient.Degree = xfFill.Gradient.Degree;
                Fill.Gradient.Top = xfFill.Gradient.Top;
                Fill.Gradient.Bottom = xfFill.Gradient.Bottom;
                Fill.Gradient.Left = xfFill.Gradient.Left;
                Fill.Gradient.Right = xfFill.Gradient.Right;
            }
            if (dxfFill != null)
            {
                if (dxfFill.PatternType != null)
                    Fill.PatternType = (ExcelFillStyle)dxfFill.PatternType;

                if (dxfFill.BackgroundColor?.Color != null)
                    Fill.BackgroundColor.SetColor(PdfColor.SetColorFromHex(dxfFill.BackgroundColor.LookupColor()));

                if (dxfFill.PatternColor?.Color != null)
                    Fill.PatternColor.SetColor(PdfColor.SetColorFromHex(dxfFill.PatternColor.LookupColor()));

                if (dxfFill.Gradient != null && dxfFill.Gradient.HasValue)
                {
                    if (dxfFill.Gradient.GradientType != null)
                        Fill.Gradient.Type = (ExcelFillGradientType)dxfFill.Gradient.GradientType;

                    if (dxfFill.Gradient.Colors[0] != null)
                        Fill.Gradient.Color1.SetColor(PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[0].Color.LookupColor()));
                    if (dxfFill.Gradient.Colors[1] != null)
                        Fill.Gradient.Color1.SetColor(PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[1].Color.LookupColor()));
                    if (dxfFill.Gradient.Colors[2] != null)
                        Fill.Gradient.Color1.SetColor(PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[2].Color.LookupColor()));
                    if (dxfFill.Gradient.Degree.HasValue && !double.IsNaN(dxfFill.Gradient.Degree.Value))
                    {
                        Fill.Gradient.Degree = (double)dxfFill.Gradient.Degree;
                    }
                    if (dxfFill.Gradient.Top.HasValue && !double.IsNaN(dxfFill.Gradient.Top.Value))
                    {
                        Fill.Gradient.Top = (double)dxfFill.Gradient.Top;
                    }
                    if (dxfFill.Gradient.Bottom.HasValue && !double.IsNaN(dxfFill.Gradient.Bottom.Value))
                    {
                        Fill.Gradient.Bottom = (double)dxfFill.Gradient.Bottom;
                    }
                    if (dxfFill.Gradient.Left.HasValue && !double.IsNaN(dxfFill.Gradient.Left.Value))
                    {
                        Fill.Gradient.Left = (double)dxfFill.Gradient.Left;
                    }
                    if (dxfFill.Gradient.Right.HasValue && !double.IsNaN(dxfFill.Gradient.Right.Value))
                    {
                        Fill.Gradient.Right = (double)dxfFill.Gradient.Right;
                    }
                }
            }
            return Fill;
        }
    }
}
