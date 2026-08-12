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
using EPPlus.Export.Pdf.Helpers;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;

namespace EPPlus.Export.Pdf.Layout
{
    [DebuggerDisplay("Cell: {Name}")]
    internal class PdfCellLayout : Transform
    {
        public PdfCellFillData CellFillData;
        public double LeftTextSpillLength = 0d;
        public double RightTextSpillLength = 0d;
        public bool delete = false;
        public bool IsHeading = false;
        public bool IsPrintTitle = false;

        public PdfCellLayout(double x, double y, double width, double height, double scaleX = 1, double scaleY = 1, double rotation = 0, Transform parent = null)
            : base(x, y - height, width, height, scaleX, scaleY, rotation, parent)
        {
            Z = 1;
            CellFillData = new PdfCellFillData();
        }

        internal void SetFill(Color backgroundColor)
        {
            CellFillData.PatternStyle = ExcelFillStyle.Solid;
            CellFillData.BackgroundColor = backgroundColor;
        }
        internal void SetPattern(PdfDictionaries dictionaries, ExcelFillStyle patternStyle, Color backgroundColor, Color patternColor)
        {
            CellFillData.PatternStyle = patternStyle;
            CellFillData.BackgroundColor = backgroundColor;
            CellFillData.PatternColor = patternColor;
            CellFillData.id = AddPatternResourceData(dictionaries.Patterns, CellFillData.PatternStyle.ToString() + CellFillData.PatternColor.ToHexString() + CellFillData.BackgroundColor.ToHexString());
        }
        internal void SetGradient(PdfDictionaries dictionaries, ExcelFillGradientType gradientType, Color color1, Color color2, Color color3, double degree, double top, double bottom, double left, double right)
        {
            if (Size.X <= 0d || Size.Y <= 0d)
            {
                return;
            }
            CellFillData.GradientFillData = new PdfCellGradientFillData();
            CellFillData.GradientFillData.GradientType = gradientType;
            CellFillData.GradientFillData.Color1 = color1;
            CellFillData.GradientFillData.Color2 = color2;
            CellFillData.GradientFillData.Color3 = color3;
            CellFillData.GradientFillData.Degree = degree;
            CellFillData.GradientFillData.Top = top;
            CellFillData.GradientFillData.Bottom = bottom;
            CellFillData.GradientFillData.Left = left;
            CellFillData.GradientFillData.Right = right;
            CellFillData.id = CellFillData.GradientFillData.ToString() + $"_{Position.X:F4}_{(Position.Y - Size.Y):F4}_{Size.X:F4}_{Size.Y:F4}";
            AddShadingResourceData(dictionaries.Shadings, CellFillData.id);
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

        public void UpdateShadingPositionMatrix(PdfPageSettings pageSettings)
        {
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
                }
            }
        }
    }
}
