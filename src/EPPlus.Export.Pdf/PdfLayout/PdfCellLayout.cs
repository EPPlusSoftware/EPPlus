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
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Export.Pdf.Pdfhelpers;

namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfCellLayout : Transform, ILayout
    {
        public PdfCellFillData CellFillData;

        public PdfCellLayout(PdfDictionaries dictionaries, ExcelRangeBase cell, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y-height, width, height, scaleX, scaleY, rotation, parent)
        {
            if (cell != null)
            {
                var fill = cell.Style.Fill;
                //Solid Fill
                CellFillData = new PdfCellFillData();
                if (fill.PatternType == ExcelFillStyle.Solid)
                {
                    var bkgc = fill.BackgroundColor;
                    CellFillData.PattenStyle = fill.PatternType;
                    if (string.IsNullOrEmpty(bkgc.LookupColor()) && !string.IsNullOrEmpty(cell.Text))
                    {
                        CellFillData.BackgroundColor = Color.Empty;
                    }
                    else
                    {
                        CellFillData.BackgroundColor = Pdfhelpers.PdfColor.SetColorFromHex(bkgc.LookupColor());
                    }
                }
                else if (fill.PatternType != ExcelFillStyle.None)
                {
                    CellFillData.PattenStyle = fill.PatternType;
                    CellFillData.BackgroundColor = PdfColor.SetColorFromHex(fill.PatternColor.Rgb == null ? "#FFFFFFFF" : fill.PatternColor.LookupColor());
                    CellFillData.PatternColor = PdfColor.SetColorFromHex(fill.BackgroundColor.LookupColor());
                    CellFillData.id = AddPatternResourceData(dictionaries.Patterns, CellFillData.PattenStyle.ToString() + CellFillData.PatternColor.ToHexString() + CellFillData.BackgroundColor.ToHexString());
                }
                else if (fill.HasGradient)
                {
                    CellFillData.GradientFillData = new PdfCellGradientFillData();
                    CellFillData.GradientFillData.GradientType = fill.Gradient.Type;
                    CellFillData.GradientFillData.Color1 = PdfColor.SetColorFromHex(fill.Gradient.Color1.LookupColor());
                    CellFillData.GradientFillData.Color2 = PdfColor.SetColorFromHex(fill.Gradient.Color2.LookupColor());
                    CellFillData.GradientFillData.Color3 = PdfColor.SetColorFromHex(fill.Gradient.Color3.LookupColor());
                    CellFillData.GradientFillData.Degree = fill.Gradient.Degree;
                    CellFillData.GradientFillData.Top = double.IsNaN(fill.Gradient.Top) ? 0 : fill.Gradient.Top;
                    CellFillData.GradientFillData.Bottom = double.IsNaN(fill.Gradient.Bottom) ? 0 : fill.Gradient.Bottom;
                    CellFillData.GradientFillData.Left = double.IsNaN(fill.Gradient.Left) ? 0 : fill.Gradient.Left;
                    CellFillData.GradientFillData.Right = double.IsNaN(fill.Gradient.Right) ? 0 : fill.Gradient.Right;
                    CellFillData.id = CellFillData.GradientFillData.ToString();
                    AddShadingResourceData(dictionaries.Shadings, CellFillData.id);

                    if (CellFillData.GradientFillData != null)
                    {
                        if (CellFillData.GradientFillData.GradientType == ExcelFillGradientType.Linear)
                        {
                            switch (CellFillData.GradientFillData.Degree)
                            {
                                case 45d:
                                    CellFillData.GradientFillData.coords = [0, 1, 1, 0];
                                    break;
                                case 90d:
                                    CellFillData.GradientFillData.coords = [0, 1, 0, 0];
                                    break;
                                case 135d:
                                    CellFillData.GradientFillData.coords = [1, 1, 0, 0];
                                    break;
                                case 180d:
                                    CellFillData.GradientFillData.coords = [1, 0, 0, 0];
                                    break;
                                case 225d:
                                    CellFillData.GradientFillData.coords = [1, 0, 0, 1];
                                    break;
                                case 270d:
                                    CellFillData.GradientFillData.coords = [0, 0, 0, 1];
                                    break;
                                case 315d:
                                    CellFillData.GradientFillData.coords = [0, 0, 1, 1];
                                    break;
                                case 0d:
                                default:
                                    CellFillData.GradientFillData.coords = [0, 0, 1, 0];
                                    break;
                            }
                            CellFillData.GradientFillData.matrix = [Size.X, 0, 0, Size.Y, LocalPosition.X, LocalPosition.Y];
                        }
                        else if (CellFillData.GradientFillData.GradientType == ExcelFillGradientType.Path)
                        {
                            //double x = LocalPosition.X;
                            //double y = LocalPosition.Y;
                            //double width = Size.X;
                            //double height = Size.Y;
                            var top = CellFillData.GradientFillData.Top;
                            var bottom = CellFillData.GradientFillData.Bottom;
                            var left = CellFillData.GradientFillData.Left;
                            var right = CellFillData.GradientFillData.Right;
                            double r = 1;
                            if (top == 0 && bottom == 0 && left == 0 && right == 0)
                            {
                                CellFillData.GradientFillData.coords = [0, 1, 0, 0, 1, r];
                            }
                            else if (top == 0 && bottom == 0 && left == 1 && right == 1)
                            {
                                CellFillData.GradientFillData.coords = [1, 1, 0, 1, 1, r];
                            }
                            else if (top == 1 && bottom == 1 && left == 0 && right == 0)
                            {
                                CellFillData.GradientFillData.coords = [0, 0, 0, 0, 0, r];
                            }
                            else if (top == 1 && bottom == 1 && left == 1 && right == 1)
                            {
                                CellFillData.GradientFillData.coords = [1, 0, 0, 1, 0, r];
                            }
                            else if (top == 0.5 && bottom == 0.5 && left == 0.5 && right == 0.5)
                            {
                                CellFillData.GradientFillData.coords = [0.5, 0.5, 0, 0.5, 0.5, r];
                            }
                            CellFillData.GradientFillData.matrix = [Size.X, 0, 0, Size.Y, LocalPosition.X, LocalPosition.Y];
                            //double cx = x;
                            //double cy = y + height;
                            //if (!double.IsNaN(left) && !double.IsNaN(right) && !double.IsNaN(top) && !double.IsNaN(bottom)) // bottom-right
                            //{
                            //    if (top < 1d && bottom < 1d &&left < 1d && right < 1d)
                            //    {
                            //        cx = x + width / 2d;
                            //        cy = y + height / 2d;
                            //    }
                            //    else
                            //    {
                            //        cx = x + width;
                            //        cy = y;
                            //    }
                            //}
                            //else if (!double.IsNaN(left) && !double.IsNaN(right)) // bottom-left
                            //{
                            //    cx = x + width;
                            //    cy = y + height;
                            //}
                            //else if (!double.IsNaN(top) && !double.IsNaN(bottom)) // top-right
                            //{
                            //    cx = x;
                            //    cy = y;
                            //}
                            //double r = System.Math.Sqrt(width * width + height * height) / 2d;
                            //CellFillData.GradientFillData.coords = [cx, cy, 0, cx, cy, r];
                        }
                    }
                }
            }
        }

        private string AddPatternResourceData(Dictionary<string, PdfPatternResource> patternResources, string key)
        {
            if (!patternResources.ContainsKey(key))
            {
                int label = 1;
                if (patternResources.Count > 0)
                {
                    label = patternResources.Last().Value.labelNumber + 1;
                }
                var pr = new PdfPatternResource(label, CellFillData);
                patternResources.Add(key, pr);
            }
            return key;
        }

        private string AddShadingResourceData(Dictionary<string, PdfShadingResource> shadingResources, string key)
        {
            if (!shadingResources.ContainsKey(key))
            {
                int label = 1;
                if (shadingResources.Count > 0)
                {
                    label = shadingResources.Last().Value.labelNumber + 1;
                }
                var pr = new PdfShadingResource(label, CellFillData);
                shadingResources.Add(key, pr);
            }
            return key;
        }

        //Adjust size and position slightly for aesthetics.
        public void AdjustForGridLines()
        {
            Size = new Vector2(Size.X + GridLine.HalfWidth, Size.Y + GridLine.HalfWidth);
            LocalPosition = new Vector2(LocalPosition.X + GridLine.FourthWidth, LocalPosition.Y + GridLine.FourthWidth);
        }

        public void ConvertCoordinates(PdfPageSettings pageSettings)
        {
            LocalPosition = new Vector2(LocalPosition.X, pageSettings.PageSize.HeightPu - System.Math.Abs(LocalPosition.Y) - Size.Y);
            if (CellFillData.GradientFillData != null)
            {
                if (CellFillData.GradientFillData.GradientType == ExcelFillGradientType.Linear)
                {
                    switch (CellFillData.GradientFillData.Degree)
                    {
                        case 45d:
                            CellFillData.GradientFillData.coords = [0, 1, 1, 0];
                            break;
                        case 90d:
                            CellFillData.GradientFillData.coords = [0, 1, 0, 0];
                            break;
                        case 135d:
                            CellFillData.GradientFillData.coords = [1, 1, 0, 0];
                            break;
                        case 180d:
                            CellFillData.GradientFillData.coords = [1, 0, 0, 0];
                            break;
                        case 225d:
                            CellFillData.GradientFillData.coords = [1, 0, 0, 1];
                            break;
                        case 270d:
                            CellFillData.GradientFillData.coords = [0, 0, 0, 1];
                            break;
                        case 315d:
                            CellFillData.GradientFillData.coords = [0, 0, 1, 1];
                            break;
                        case 0d:
                        default:
                            CellFillData.GradientFillData.coords = [0, 0, 1, 0];
                            break;
                    }
                    CellFillData.GradientFillData.matrix = [Size.X, 0, 0, Size.Y, LocalPosition.X, LocalPosition.Y];
                }
                else if (CellFillData.GradientFillData.GradientType == ExcelFillGradientType.Path)
                {
                    double x = LocalPosition.X;
                    double y = LocalPosition.Y;
                    double width = Size.X;
                    double height = Size.Y;
                    var top = CellFillData.GradientFillData.Top;
                    var bottom = CellFillData.GradientFillData.Bottom;
                    var left = CellFillData.GradientFillData.Left;
                    var right = CellFillData.GradientFillData.Right;
                    double r = 1;
                    if (top == 0 && bottom == 0 && left == 0 && right == 0)
                    {
                        CellFillData.GradientFillData.coords = [0, 1, 0, 0, 1, r];
                    }
                    else if (top == 0 && bottom == 0 && left == 1 && right == 1)
                    {
                        CellFillData.GradientFillData.coords = [1, 1, 0, 1, 1, r];
                    }
                    else if (top == 1 && bottom == 1 && left == 0 && right == 0)
                    {
                        CellFillData.GradientFillData.coords = [0, 0, 0, 0, 0, r];
                    }
                    else if (top == 1 && bottom == 1 && left == 1 && right == 1)
                    {
                        CellFillData.GradientFillData.coords = [1, 0, 0, 1, 0, r];
                    }
                    else if (top == 0.5 && bottom == 0.5 && left == 0.5 && right == 0.5)
                    {
                        CellFillData.GradientFillData.coords = [0.5, 0.5, 0, 0.5, 0.5, r];
                    }
                    CellFillData.GradientFillData.matrix = [Size.X, 0, 0, Size.Y, LocalPosition.X, LocalPosition.Y];
                    //double cx = x;
                    //double cy = y + height;
                    //if (!double.IsNaN(left) && !double.IsNaN(right) && !double.IsNaN(top) && !double.IsNaN(bottom)) // bottom-right
                    //{
                    //    if (top < 1d && bottom < 1d &&left < 1d && right < 1d)
                    //    {
                    //        cx = x + width / 2d;
                    //        cy = y + height / 2d;
                    //    }
                    //    else
                    //    {
                    //        cx = x + width;
                    //        cy = y;
                    //    }
                    //}
                    //else if (!double.IsNaN(left) && !double.IsNaN(right)) // bottom-left
                    //{
                    //    cx = x + width;
                    //    cy = y + height;
                    //}
                    //else if (!double.IsNaN(top) && !double.IsNaN(bottom)) // top-right
                    //{
                    //    cx = x;
                    //    cy = y;
                    //}
                    //double r = System.Math.Sqrt(width * width + height * height) / 2d;
                    //CellFillData.GradientFillData.coords = [cx, cy, 0, cx, cy, r];
                }
            }
        }
    }
}