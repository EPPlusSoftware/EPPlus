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
                }
                else
                {
                    try
                    {
                        if (fill.Gradient != null && fill.Gradient.Type != Style.ExcelFillGradientType.None)
                        {
                            CellFillData.GradientFillData = new PdfCellGradientFillData();
                            CellFillData.GradientFillData.GradientType = fill.Gradient.Type;
                            CellFillData.GradientFillData.Color0 = new PdfColor(fill.Gradient.Color1.LookupColor());
                            CellFillData.GradientFillData.Color1 = new PdfColor(fill.Gradient.Color2.LookupColor());
                            CellFillData.GradientFillData.Degree = double.IsNaN(fill.Gradient.Degree) ? 0d : fill.Gradient.Degree;
                            CellFillData.GradientFillData.Top = double.IsNaN(fill.Gradient.Top) ? y : fill.Gradient.Top;
                            CellFillData.GradientFillData.Bottom = double.IsNaN(fill.Gradient.Bottom) ? y - height : fill.Gradient.Bottom;
                            CellFillData.GradientFillData.Left = double.IsNaN(fill.Gradient.Left) ? x : fill.Gradient.Left;
                            CellFillData.GradientFillData.Right = double.IsNaN(fill.Gradient.Right) ? x + width : fill.Gradient.Right;
                            //CellFillData.GradientFillData.matrix = [width, 0d, 0d, height, x, y];
                            //CellFillData.GradientFillData.id = AddPatternResourceData(dictionaries.Patterns, CellFillData.GradientFillData.ToString());
                            CellFillData.GradientFillData.id = AddShadingResourceData(dictionaries.Shadings, CellFillData.GradientFillData.ToString());
                        }
                    }
                    catch(InvalidCastException)
                    {

                    }
                }
                //Pattern Fill
                CellFillData.PattenStyle = cell.Style.Fill.PatternType;
                CellFillData.PatternColor = new PdfColor(cell.Style.Fill.PatternColor.LookupColor());
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
                var rad = -CellFillData.GradientFillData.Degree * System.Math.PI / 180d;
                double cos = System.Math.Cos(rad);
                double sin = System.Math.Sin(rad);
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
                double cx = LocalPosition.X + Size.X / 2d;
                double cy = LocalPosition.Y + Size.Y / 2d;

                // Half-diagonal vector along the gradient axis
                // We use max(w,h) to ensure gradient fully covers rectangle
                double half = System.Math.Sqrt(Size.X * Size.X + Size.Y * Size.Y) / 2d;

                double dx = cos * half;
                double dy = sin * half;

                // Coords
                double x0 = cx - dx;
                double y0 = cy - dy;
                double x1 = cx + dx;
                double y1 = cy + dy;
                CellFillData.GradientFillData.coords = [x0, y0, x1, y1];
            }
        }
    }
}
