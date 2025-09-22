using OfficeOpenXml.PDF.Math;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.PdfResources;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeOpenXml.PDF.PdfLayout
{
    internal class PdfCellLayout : PdfTransform
    {
        public PdfCellFillData CellFillData;

        public PdfCellLayout(Dictionary<string, PdfPatternResource> patternResources, ExcelRangeBase cell, double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, PdfTransform parent = null)
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
                            CellFillData.GradientFillData.id = AddPatternResourceData(patternResources, CellFillData.GradientFillData.ToString());
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
                var pr = new PdfPatternResource(label, CellFillData.GradientFillData);
                patternResources.Add(key, pr);
            }
            return key;
        }

        //Adjust size and position slightly for aesthetics.
        public void AdjustForGridLines()
        {
            Size = new Vector2(Size.X + GridLine.HalfWidth, Size.Y + GridLine.HalfWidth);
            LocalPosition = new Vector2(LocalPosition.X + GridLine.FourthWidth, LocalPosition.Y + GridLine.HalfWidth + GridLine.FourthWidth);
        }
    }
}
