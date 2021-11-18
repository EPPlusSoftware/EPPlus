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
                    svg = string.Format(PatternFills.Dott75, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.MediumGray:
                    svg = string.Format(PatternFills.Dott50, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.LightGray:
                    svg = string.Format(PatternFills.Dott25, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.Gray125:
                    svg = string.Format(PatternFills.Dott12_5, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.Gray0625:
                    svg = string.Format(PatternFills.Dott6_25, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.DarkHorizontal:
                    svg = string.Format(PatternFills.HorizontalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.DarkVertical:
                    svg = string.Format(PatternFills.VerticalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.LightHorizontal:
                    svg = string.Format(PatternFills.ThinHorizontalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.LightVertical:
                    svg = string.Format(PatternFills.ThinVerticalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.DarkDown:
                    svg = string.Format(PatternFills.ReverseDiagonalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.DarkUp:
                    svg = string.Format(PatternFills.DiagonalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.LightDown:
                    svg = string.Format(PatternFills.ThinReverseDiagonalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.LightUp:
                    svg = string.Format(PatternFills.ThinDiagonalStripe, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.DarkGrid:
                    svg = string.Format(PatternFills.DiagonalCrosshatch, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.DarkTrellis:
                    svg = string.Format(PatternFills.ThickDiagonalCrosshatch, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.LightGrid:
                    svg = string.Format(PatternFills.ThinHorizontalCrosshatch, backgroundColor, patternColor);
                    break;
                case ExcelFillStyle.LightTrellis:
                    svg = string.Format(PatternFills.ThinDiagonalCrosshatch, backgroundColor, patternColor);
                    break;
                default:
                    return "";
            }

            return $"background-repeat:repeat;background:url(data:image/svg+xml;base64,{Convert.ToBase64String(Encoding.ASCII.GetBytes(svg))});";
        }

    }
}
