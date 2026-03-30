using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FillDrawingBasic : IFill
    {
        ExcelDrawingFillBasic _fill;

        internal FillDrawingBasic(ExcelDrawingFillBasic fill)
        {
            _fill = fill;
        }

        public bool IsGradient
        {
            get
            {
                return _fill.Style == eFillStyle.GradientFill;
            }
        }


        public bool IsLinear
        {
            get
            {
                return _fill.GradientFill.ShadePath == eShadePath.Linear;
            }
        }


        public double Degree
        {
            get
            {
                if (IsGradient)
                {
                    if (IsLinear)
                    {
                        return _fill.GradientFill.LinearSettings.Angle;
                    }
                }

                return double.NaN;
            }
        }



        public double Right
        {
            get
            {
                if (IsGradient)
                {
                    if (IsLinear == false)
                    {
                        return _fill.GradientFill.TileRectangle.RightOffset;
                    }
                }

                return double.NaN;
            }
        }

        public double Bottom
        {
            get
            {
                if (IsGradient)
                {
                    if (IsLinear == false)
                    {
                        return _fill.GradientFill.TileRectangle.BottomOffset;
                    }
                }

                return double.NaN;
            }
        }

        public string GetBackgroundColor(ExcelTheme theme)
        {
            return GetColor(_fill.Color, theme);
        }

        public string GetPatternColor(ExcelTheme theme)
        {
            return GetColor(Color.Empty, theme);
        }

        public string GetGradientColor1(ExcelTheme theme)
        {
            return GetColor(_fill.GradientFill.Colors.ToArray()[0].Color.GetColor(), theme);
        }
        public string GetGradientColor2(ExcelTheme theme)
        {
            return GetColor(_fill.GradientFill.Colors.ToArray()[1].Color.GetColor(), theme);
        }

        public bool HasValue
        {
            get
            {
                return !_fill.IsEmpty;
            }
        }

        public ExcelFillStyle PatternType => ExcelFillStyle.Solid;

        internal static string GetColor(Color c, ExcelTheme theme)
        {
            return "#" + c.ToArgb().ToString("x8").Substring(2);
        }

        string IFill.GetBackgroundColor(ExcelTheme theme)
        {
            return GetBackgroundColor(theme);
        }

        string IFill.GetPatternColor(ExcelTheme theme)
        {
            return GetPatternColor(theme);
        }

        string IFill.GetGradientColor1(ExcelTheme theme)
        {
            return GetGradientColor1(theme);
        }

        string IFill.GetGradientColor2(ExcelTheme theme)
        {
            return GetGradientColor2(theme);
        }
    }
}
