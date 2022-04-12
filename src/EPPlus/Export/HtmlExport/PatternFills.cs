/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    internal class PatternFills
    {
        internal const string Dott75 =                    "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='2'><rect width='4' height='2' fill='{1}'/><rect x='0' y='0' width='1' height='1' fill='{0}'/><rect x='2' y='1' width='1' height='1' fill='{0}'/></svg>";
        internal const string Dott50 =                    "<svg xmlns='http://www.w3.org/2000/svg' width='2' height='2'><rect width='2' height='2' fill='{0}'/><rect x='0' y='0' width='1' height='1' fill='{1}'/><rect x='1' y='1' width='1' height='1' fill='{1}'/></svg>";
        internal const string Dott25 =                    "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='2'><rect width='4' height='2' fill='{0}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='0' y='1' width='1' height='1' fill='{1}'/></svg>";
        internal const string Dott12_5 =                  "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='1' y='3' width='1' height='1' fill='{1}'/></svg>";
        internal const string Dott6_25 =                  "<svg xmlns='http://www.w3.org/2000/svg' width='8' height='4'><rect width='8' height='4' fill='{0}'/><rect x='7' y='0' width='1' height='1' fill='{1}'/><rect x='3' y='2' width='1' height='1' fill='{1}' /></svg>";
        internal const string HorizontalStripe =          "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='4'><rect width='1' height='4' fill='{0}'/><rect x='0' y='1' width='1' height='2' fill='{1}'/></svg>";
        internal const string VerticalStripe =            "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='1'><rect width='4' height='1' fill='{0}'/><rect x='1' y='0' width='2' height='2' fill='{1}'/></svg>";
        internal const string ThinHorizontalStripe =      "<svg xmlns='http://www.w3.org/2000/svg' width='1' height='4'><rect width='1' height='4' fill='{0}'/><rect x='0' y='1' width='1' height='1' fill='{1}'/></svg>";
        internal const string ThinVerticalStripe =        "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='1'><rect width='4' height='1' fill='{0}'/><rect x='1' y='0' width='2' height='1' fill='{1}'/></svg>";
                                                         
        internal const string ReverseDiagonalStripe =     "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='0' y='1' width='1' height='1' fill='{1}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='2' height='1' fill='{1}'/><rect x='1' y='3' width='2' height='1' fill='{1}'/></svg>";
        internal const string DiagonalStripe =            "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='1' y='1' width='2' height='1' fill='{1}'/><rect x='0' y='2' width='2' height='1' fill='{1}'/><rect x='0' y='3' width='1' height='1' fill='{1}'/><rect x='3' y='3' width='1' height='1' fill='{1}'/></svg>";

        internal const string ThinReverseDiagonalStripe = "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='1' height='1' fill='{1}'/><rect x='1' y='3' width='1' height='1' fill='{1}'/></svg>";
        internal const string ThinDiagonalStripe =        "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='1' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='1' height='1' fill='{1}'/><rect x='3' y='3' width='1' height='1' fill='{1}'/></svg>";
        
        internal const string DiagonalCrosshatch =        "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='2' y='0' width='2' height='2' fill='{1}'/><rect x='0' y='2' width='2' height='2' fill='{1}'/></svg>";                
        internal const string ThickDiagonalCrosshatch =   "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='2' y='0' width='2' height='1' fill='{1}'/><rect x='0' y='1' width='4' height='1' fill='{1}'/><rect x='0' y='2' width='2' height='1' fill='{1}'/><rect x='0' y='3' width='4' height='1' fill='{1}'/></svg>";
        internal const string ThinHorizontalCrosshatch =  "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='3' y='0' width='1' height='1' fill='{1}'/><rect x='0' y='1' width='4' height='1' fill='{1}'/><rect x='3' y='2' width='1' height='1' fill='{1}'/><rect x='3' y='3' width='1' height='1' fill='{1}'/></svg>";        
        internal const string ThinDiagonalCrosshatch =    "<svg xmlns='http://www.w3.org/2000/svg' width='4' height='4'><rect width='4' height='4' fill='{0}'/><rect x='0' y='0' width='1' height='1' fill='{1}'/><rect x='2' y='0' width='1' height='1' fill='{1}'/><rect x='3' y='1' width='1' height='1' fill='{1}'/><rect x='0' y='2' width='1' height='1' fill='{1}'/><rect x='2' y='2' width='1' height='1' fill='{1}'/><rect x='1' y='3' width='1' height='1' fill='{1}'/></svg>";

        internal static string GetPatternSvg(ExcelFillStyle patternType, string backgroundColor, string patternColor)
        {
            string svg;
            switch (patternType)
            {
                case ExcelFillStyle.DarkGray:
                    svg = string.Format(Dott75, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.MediumGray:
                    svg = string.Format(Dott50, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.LightGray:
                    svg = string.Format(Dott25, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.Gray125:
                    svg = string.Format(Dott12_5, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.Gray0625:
                    svg = string.Format(Dott6_25, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.DarkHorizontal:
                    svg = string.Format(HorizontalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.DarkVertical:
                    svg = string.Format(VerticalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.LightHorizontal:
                    svg = string.Format(ThinHorizontalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.LightVertical:
                    svg = string.Format(ThinVerticalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.DarkDown:
                    svg = string.Format(ReverseDiagonalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.DarkUp:
                    svg = string.Format(DiagonalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.LightDown:
                    svg = string.Format(ThinReverseDiagonalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.LightUp:
                    svg = string.Format(ThinDiagonalStripe, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.DarkGrid:
                    svg = string.Format(DiagonalCrosshatch, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.DarkTrellis:
                    svg = string.Format(ThickDiagonalCrosshatch, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.LightGrid:
                    svg = string.Format(ThinHorizontalCrosshatch, patternColor, backgroundColor);
                    break;
                case ExcelFillStyle.LightTrellis:
                    svg = string.Format(ThinDiagonalCrosshatch, patternColor, backgroundColor);
                    break;
                default:
                    return "";
            }

            return $"background-repeat:repeat;background:url(data:image/svg+xml;base64,{Convert.ToBase64String(Encoding.ASCII.GetBytes(svg))});";
        }

    }
}
