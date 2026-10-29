using EPPlus.Export.Pdf.Helpers;
using EPPlus.Export.Pdf.Layout;
using EPPlus.Export.Pdf.Resources;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfFillResolver
    {
        internal static void SetFill(PdfDictionaries dictionaries, PdfCellStyle cellStyle, string text, PdfCellLayout fill)
        {
            var xfFill = cellStyle.xfFill;
            var dxfFill = cellStyle.dxfFill;
            if (dxfFill != null && xfFill.IsEmpty())
            {
                var patternStyle = dxfFill.PatternType != null ? (ExcelFillStyle)dxfFill.PatternType : ExcelFillStyle.Solid;
                if (patternStyle == ExcelFillStyle.Solid)
                {
                    fill.SetFill(PdfColor.SetColorFromHex(dxfFill.BackgroundColor.LookupColor()));
                }
                else if (patternStyle != ExcelFillStyle.None)
                {
                    var bkgc = PdfColor.SetColorFromHex(dxfFill.PatternColor.Color == null ? "#FFFFFFFF" : dxfFill.PatternColor.LookupColor());
                    var patc = PdfColor.SetColorFromHex(dxfFill.BackgroundColor.LookupColor());
                    fill.SetPattern(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillStyle)patternStyle, bkgc, patc);
                }
                else if (dxfFill.Gradient != null)
                {
                    var gradientType = dxfFill.Gradient.GradientType == null ? ExcelFillGradientType.None : (ExcelFillGradientType)dxfFill.Gradient.GradientType;
                    var color1 = PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[0].Color.LookupColor());
                    var color2 = PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[1].Color.LookupColor());
                    var color3 = PdfColor.SetColorFromHex(dxfFill.Gradient.Colors[2].Color.LookupColor());
                    var degree = dxfFill.Gradient.Degree == null ? 0 : (double)dxfFill.Gradient.Degree;
                    var top = dxfFill.Gradient.Top == null ? 0 : (double)dxfFill.Gradient.Top;
                    var bottom = dxfFill.Gradient.Bottom == null ? 0 : (double)dxfFill.Gradient.Bottom;
                    var left = dxfFill.Gradient.Left == null ? 0 : (double)dxfFill.Gradient.Left;
                    var right = dxfFill.Gradient.Right == null ? 0 : (double)dxfFill.Gradient.Right;
                    fill.SetGradient(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillGradientType)gradientType, color1, color2, color3, degree, top, bottom, left, right);
                }
            }
            else
            {
                if (xfFill.PatternType == ExcelFillStyle.Solid)
                {
                    var bkgc = xfFill.BackgroundColor;
                    var patternStyle = xfFill.PatternType;
                    if (string.IsNullOrEmpty(bkgc.LookupColor()) && !string.IsNullOrEmpty(text))
                    {
                        fill.SetFill(Color.Empty);
                    }
                    else
                    {
                        fill.SetFill(PdfColor.SetColorFromHex(bkgc.LookupColor()));
                    }
                }
                else if (xfFill.PatternType != ExcelFillStyle.None)
                {
                    var patternStyle = xfFill.PatternType;
                    var bkgc = PdfColor.SetColorFromHex(xfFill.PatternColor.Rgb == null ? "#FFFFFFFF" : xfFill.PatternColor.LookupColor());
                    var patc = PdfColor.SetColorFromHex(xfFill.BackgroundColor.LookupColor());
                    fill.SetPattern(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillStyle)patternStyle, bkgc, patc);
                }
                else if (xfFill.HasGradient)
                {
                    var gradientType = xfFill.Gradient.Type;
                    var color1 = PdfColor.SetColorFromHex(xfFill.Gradient.Color1.LookupColor());
                    var color2 = PdfColor.SetColorFromHex(xfFill.Gradient.Color2.LookupColor());
                    var color3 = PdfColor.SetColorFromHex(xfFill.Gradient.Color3.LookupColor());
                    var degree = xfFill.Gradient.Degree;
                    var top = double.IsNaN(xfFill.Gradient.Top) ? 0 : xfFill.Gradient.Top;
                    var bottom = double.IsNaN(xfFill.Gradient.Bottom) ? 0 : xfFill.Gradient.Bottom;
                    var left = double.IsNaN(xfFill.Gradient.Left) ? 0 : xfFill.Gradient.Left;
                    var right = double.IsNaN(xfFill.Gradient.Right) ? 0 : xfFill.Gradient.Right;
                    fill.SetGradient(dictionaries, (EPPlus.Export.Pdf.Enums.ExcelFillGradientType)gradientType, color1, color2, color3, degree, top, bottom, left, right);
                }
            }
        }
    }
}
