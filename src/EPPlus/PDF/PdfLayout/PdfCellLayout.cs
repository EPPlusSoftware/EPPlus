using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.PdfResources;
using OfficeOpenXml.PDF.PdfSettings;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellLayout : PdfTransform, ILayout
    {
        public PdfCellFillData CellFillData;

        public PdfCellLayout(PdfDictionaries dictionaries, ExcelRangeBase cell, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
            :base(x, y, width, height, scaleX, scaleY, rotation, parent )
        {
            if (cell != null)
            {
                var fill = cell.Style.Fill;
                //Solid Fill
                CellFillData = new PdfCellFillData();
                if (fill.PatternType == Style.ExcelFillStyle.Solid)
                {
                    var bkgc = fill.BackgroundColor;
                    if (string.IsNullOrEmpty(bkgc.LookupColor()) && !string.IsNullOrEmpty(cell.Text))
                    {
                        CellFillData.BackgroundColor = PdfColor.None;
                    }
                    else
                    {
                        CellFillData.BackgroundColor = new PdfColor(bkgc.LookupColor());
                    }
                }
                else if (fill.PatternType != Style.ExcelFillStyle.None)
                {
                    CellFillData.PattenStyle = fill.PatternType;
                    CellFillData.PatternColor = new PdfColor(fill.PatternColor.LookupColor());
                    CellFillData.BackgroundColor = new PdfColor(fill.BackgroundColor.LookupColor());
                }
                else
                {
                    try
                    {
                        if (fill.Gradient != null && fill.Gradient.Type != Style.ExcelFillGradientType.None)
                        {
                            CellFillData.GradientFillData = new PdfCellGradientFillData();
                            CellFillData.GradientFillData.GradientType = fill.Gradient.Type;
                            CellFillData.GradientFillData.Color1 = new PdfColor(fill.Gradient.Color1.LookupColor());
                            CellFillData.GradientFillData.Color2 = new PdfColor(fill.Gradient.Color2.LookupColor());
                            CellFillData.GradientFillData.Color3 = new PdfColor(fill.Gradient.Color3.LookupColor());
                            CellFillData.GradientFillData.Degree = double.IsNaN(fill.Gradient.Degree) ? 0d : fill.Gradient.Degree;
                            CellFillData.GradientFillData.Top = fill.Gradient.Top;
                            CellFillData.GradientFillData.Bottom = fill.Gradient.Bottom;
                            CellFillData.GradientFillData.Left = fill.Gradient.Left;
                            CellFillData.GradientFillData.Right = fill.Gradient.Right;
                            //CellFillData.GradientFillData.matrix = [width, 0d, 0d, height, x, y];
                            //CellFillData.GradientFillData.id = AddPatternResourceData(dictionaries.Patterns, CellFillData.GradientFillData.ToString());
                            CellFillData.GradientFillData.id = AddShadingResourceData(dictionaries.Shadings, CellFillData.GradientFillData.ToString());
                        }
                    }
                    catch(InvalidCastException)
                    {

                    }
                }
            }
        }

        private string AddPatternResourceData(Dictionary<string, PdfPatternResource> patternResources, string key)
        {
            if(!patternResources.ContainsKey(key))
            {
                int label = 1;
                if (patternResources.Count > 0)
                {
                    label = patternResources.Last().Value.labelNumber + 1;
                }
                var pr = new PdfPatternResource(label, CellFillData.GradientFillData); //send cell data here and calculate matrix for SHadingPattern(Need to recalculate y for cell...). Create the objects in constructor isntead of later when adding them to the document. (maybe do the same for font to keep it similar)
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
                var pr = new PdfShadingResource(label, CellFillData.GradientFillData);
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
            LocalPosition = new Vector2(LocalPosition.X, (pageSettings.PageSize.HeightPu - System.Math.Abs(LocalPosition.Y) - Size.Y));
            if (CellFillData.GradientFillData != null)
            {
                if (CellFillData.GradientFillData.GradientType == Style.ExcelFillGradientType.Linear)
                {
                    var rad = -CellFillData.GradientFillData.Degree * System.Math.PI / 180d;
                    double cos = System.Math.Cos(rad);
                    double sin = System.Math.Sin(rad);
                    double cx = LocalPosition.X + Size.X / 2d;
                    double cy = LocalPosition.Y + Size.Y / 2d;
                    //Setting matrix in pattern object
                    //double a = Size.X * cos;
                    //double b = Size.X * sin;
                    //double c = -Size.Y * sin;
                    //double d = Size.Y * cos;
                    //if (CellFillData.GradientFillData.Degree == 0)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X, LocalPosition.Y];
                    //else if (CellFillData.GradientFillData.Degree == 45)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X, LocalPosition.Y + Size.Y];
                    //else if (CellFillData.GradientFillData.Degree == 90)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X, LocalPosition.Y + Size.Y * 2.5];
                    //else if (CellFillData.GradientFillData.Degree == 135)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X + Size.X, LocalPosition.Y + Size.Y];
                    //else if (CellFillData.GradientFillData.Degree == 180)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X + Size.X, LocalPosition.Y + Size.Y];
                    //else if (CellFillData.GradientFillData.Degree == 225)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X + Size.X, LocalPosition.Y];
                    //else if (CellFillData.GradientFillData.Degree == 270)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X, LocalPosition.Y - (Size.Y / 2)];
                    //else if (CellFillData.GradientFillData.Degree == 315)
                    //    CellFillData.GradientFillData.matrix = [a, b, c, d, LocalPosition.X, LocalPosition.Y];
                    //Setting Coords in Shading object
                    // Midpoint of the rectangle
                    // Half-diagonal vector along the gradient axis
                    // We use max(w,h) to ensure gradient fully covers rectangle
                    //ouble half = System.Math.Sqrt(Size.X * Size.X + Size.Y * Size.Y) / 2d;
                    double projW = System.Math.Abs(Size.X * cos);
                    double projH = System.Math.Abs(Size.Y * sin);

                    // Effective half-length of the gradient axis
                    double half = (projW + projH) / 2d;

                    double dx = cos * half;
                    double dy = sin * half;

                    // Coords
                    double x0 = cx - dx;
                    double y0 = cy - dy;
                    double x1 = cx + dx;
                    double y1 = cy + dy;
                    CellFillData.GradientFillData.coords = [x0, y0, x1, y1];
                }
                else if( CellFillData.GradientFillData.GradientType == Style.ExcelFillGradientType.Path)
                {
                    double x = LocalPosition.X;
                    double y = LocalPosition.Y;
                    double width = Size.X;
                    double height = Size.Y;
                    var top = CellFillData.GradientFillData.Top;
                    var bottom = CellFillData.GradientFillData.Bottom;
                    var left = CellFillData.GradientFillData.Left;
                    var right = CellFillData.GradientFillData.Right;

                    // Default corner = top-left
                    double cx = x;
                    double cy = y + height;

                    if (!double.IsNaN(left) && !double.IsNaN(right) && !double.IsNaN(top) && !double.IsNaN(bottom)) // bottom-right
                    {
                        if (top < 1d && bottom < 1d &&left < 1d && right < 1d)
                        {
                            cx = x + width / 2d;
                            cy = y + height / 2d;
                        }
                        else
                        {
                            cx = x + width;
                            cy = y;
                        }
                    }
                    else if (!double.IsNaN(left) && !double.IsNaN(right)) // bottom-left
                    {
                        cx = x + width;
                        cy = y + height;

                    }
                    else if (!double.IsNaN(top) && !double.IsNaN(bottom)) // top-right
                    {
                        cx = x;
                        cy = y;
                    }
                    //else // default: top-left
                    //{
                    //    cx = x;
                    //    cy = y + height;
                    //}

                    // radius = diagonal length (corner to opposite corner)
                    double r = System.Math.Sqrt(width * width + height * height) / 2d;

                    CellFillData.GradientFillData.coords = [cx, cy, 0, cx, cy, r];
                }
            }
        }
    }
}
