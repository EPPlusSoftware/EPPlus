using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using tc = OfficeOpenXml.Utils.TypeConversion;

namespace OfficeOpenXml.Drawing.Renderer.Chart.ChartElementStyleTables
{
    [Flags]
    enum ChartElement
    {
        None = 0,
        ChartArea = 1,
        PlotArea2d = 2,
        PloatArea3d = 4,
        Axis = 8,
        MinorGridLines = 16,
        MajorGridLines = 32,
        DataTable = 64,
        Floor = 128,
        Walls = 256,
        OtherLines = 512,
    }

    internal class ChartDrawingObjectWithDefaults : ChartDrawingObject
    {
        public ChartDrawingObjectWithDefaults(ChartRenderer chart) : base(chart)
        {

        }

        private Color GetSchemeColorTint(eSchemeColor sColor, double tint = 0.0d)
        {
            var schemeClr = tc.ColorConverter.GetSchemeColor(ChartRenderer.Theme, sColor);
            var tintedSchemeColor = tc.ColorConverter.ApplyTintDrawing(schemeClr, tint);
            return tintedSchemeColor;
        }

        private Color GetThemeColorTint(eThemeSchemeColor themeColor, double tint = 0.0d)
        {
            var schemeClr = tc.ColorConverter.GetThemeColor(ChartRenderer.Theme, themeColor);
            var tintedSchemeColor = tc.ColorConverter.ApplyTintDrawing(schemeClr, tint);
            return tintedSchemeColor;
        }

        internal Color? GetStyleColorOrDefault(int styleId, Color col1, Color col2, Color col3, Color col4)
        {
            Color? themeColor = null;
            //Chart style can only be above 48 if it is Style102 which in this case should be equivalent with style2
            //Alternatively it's an unkown or unset style which should also default to style2
            styleId = styleId > (int)eChartStyle.Style48 ? (int)eChartStyle.Style2 : styleId;

            if (styleId == 0)
            {
                return Color.Empty;
            }

            if (styleId <= 32)
            {
                themeColor = col1;
            }
            else if (styleId <= 34)
            {
                themeColor = col2;
            }
            else if (styleId <= 40)
            {
                themeColor = col3;
            }
            else if (styleId <= 48)
            {
                themeColor = col4;
            }

            return themeColor;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <param name="ChartStyleId"></param>
        /// <param name="lineColor">The line color with fill styles etc applied</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        protected ExcelThemeLine GetThemedLine(ChartElement element, int ChartStyleId, out Color? lineColor)
        {
            if(element.HasFlag(ChartElement.Floor | ChartElement.ChartArea))
            {
                lineColor = GetDefaultBorderColorForElement(element, ChartStyleId);
                var themedLine = ChartRenderer.Theme.FormatScheme.BorderStyle[0];

                if (themedLine.HasFill == false)
                {
                    //Node exists but has no fill. Excel considers this the same as transparent/noFill
                    lineColor = Color.Transparent;
                    return themedLine;
                }

                if (ChartStyleId < 41)
                {
                    if (themedLine.Fill.SolidFill.Color.Transforms.Count > 0)
                    {
                        //themeColor = tc.ColorConverter.ApplyTintDrawing(themeColor.Value, 0.15d);
                        //but even in this case if there is no ln node found in style it appears to default to 75% despite a scheme color existing in the theme
                        lineColor = tc.ColorConverter.ApplyTransforms(lineColor.Value, themedLine.Fill.SolidFill.Color.Transforms);
                    }
                    else
                    {
                        //Default value Should arguably be 75% tint themeColor but something is strange...
                        //It appears closer to 50% in this specific case
                        //It also appears to be tx1 (black) and apply color and tint 0.25 in vba
                        var newTheme = tc.ColorConverter.ApplyTintDrawing(lineColor.Value, 0.25d);
                        lineColor = newTheme;
                    }
                }
                else
                {
                    //No Line
                    lineColor = Color.Transparent;
                    return null;
                }

                return themedLine;
            }
            else
            {
                throw new InvalidOperationException(
                    $"The enum option: '{Enum.GetName(typeof(ChartElement), element)}' is invalid. " +
                    $"Only ChartArea or Floor has a default themed line");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="element"></param>
        /// <param name="ChartStyleId"></param>
        /// <param name="fillColor">The fill color with fill styles etc applied</param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException"></exception>
        protected ExcelDrawingFill GetThemedFill(ChartElement element, int ChartStyleId, out Color? fillColor)
        {
            var bg = ChartRenderer.Theme.FormatScheme.BackgroundFillStyle[0];

            if (element.HasFlag(ChartElement.Floor | ChartElement.Walls))
            {
                fillColor = GetDefaultFillColorForElement(element, ChartStyleId);
                var themedFill = ChartRenderer.Theme.FormatScheme.BackgroundFillStyle[0];

                if (ChartStyleId > 32)
                {
                    if (themedFill.SolidFill.Color.Transforms.Count > 0)
                    {
                        fillColor = tc.ColorConverter.ApplyTransforms(fillColor.Value, themedFill.SolidFill.Color.Transforms);
                    }
                    else
                    {
                        //Hardcoded default for fills without any actual info in excel
                        var newTheme = tc.ColorConverter.ApplyTintDrawing(fillColor.Value, 0.75d);
                        fillColor = newTheme;
                    }
                }
                else
                {
                    //No Fill
                    fillColor = Color.Transparent;
                    return null;
                }

                return themedFill;
            }
            else
            {
                throw new InvalidOperationException(
                    $"The enum option: '{Enum.GetName(typeof(ChartElement), element)}' is invalid. " +
                    $"Only Walls or Floor has a default themed line");
            }
        }

        protected Color? GetDefaultBorderColorForElement(ChartElement element, int ChartStyleId)
        {
            //return Color.Empty;

            if(element.HasFlag(ChartElement.Axis | ChartElement.MajorGridLines))
            {
                //There's only really two options in this particular case
                if(ChartStyleId <= 32)
                {
                    return GetSchemeColorTint(eSchemeColor.Text1, 0.75d);
                }
                else
                {
                    return GetSchemeColorTint(eSchemeColor.Background1, 0.75d);
                }
            }
            else if(element.HasFlag(ChartElement.MinorGridLines))
            {
                var retCol = GetSchemeColorTint(eSchemeColor.Text1, 0.5d);
                var retCol2and3 = GetSchemeColorTint(eSchemeColor.Background1, 0.5d);
                var retCol4 = GetSchemeColorTint(eSchemeColor.Background1, 0.9d);

                return GetStyleColorOrDefault(ChartStyleId, retCol, retCol2and3, retCol2and3, retCol4);
            }
            else if (element.HasFlag(ChartElement.ChartArea | ChartElement.DataTable | ChartElement.Floor))
            {
                var retCol = GetSchemeColorTint(eSchemeColor.Text1, 0.75d);
                var retCol2and3 = GetSchemeColorTint(eSchemeColor.Background1, 0.75d);
                var retCol4 = GetSchemeColorTint(eSchemeColor.Text1, 1d);

                return GetStyleColorOrDefault(ChartStyleId, retCol, retCol2and3, retCol2and3, retCol4);
            }
            else
            {
                //Other lines should technically always be the enum here but keep it as Else just in case
                var retCol = GetSchemeColorTint(eSchemeColor.Text1, 1d);
                var retCol2and3 = GetSchemeColorTint(eSchemeColor.Background1, 1d);
                var retCol4 = GetSchemeColorTint(eSchemeColor.Text1, 1d);

                return GetStyleColorOrDefault(ChartStyleId, retCol, retCol2and3, retCol2and3, retCol4);
            }
        }


        private Color? GetDefaultAccent(int ChartStyleId)
        {
            if(ChartStyleId < 35 || ChartStyleId > 40)
            {
                throw new InvalidOperationException($"Invalid ChartStyleId '{ChartStyleId}'" +
                    $"Default Accent tint must be between 35 and 40");
            }
            //35 == accent1, 36 == accent2 etc.
            var accentColor = (eSchemeColor.Accent1 + (ChartStyleId) - 35);
            return GetSchemeColorTint(accentColor, 0.2d);
        }

        protected Color? GetDefaultFillColorForElement(ChartElement element, int ChartStyleId)
        {
            if (element.HasFlag(ChartElement.ChartArea))
            {
                var retCol = GetSchemeColorTint(eSchemeColor.Background1);
                var retCol2And3 = GetSchemeColorTint(eSchemeColor.Text1);
                var retCol4 = GetSchemeColorTint(eSchemeColor.Background1);

                return GetStyleColorOrDefault(ChartStyleId, retCol, retCol2And3, retCol2And3, retCol4);
            }
            else if(element.HasFlag(ChartElement.Floor | ChartElement.Walls | ChartElement.PlotArea2d))
            {
                var retCol = GetSchemeColorTint(eSchemeColor.Background1);
                var retCol2 = GetSchemeColorTint(eSchemeColor.Background1, 0.2d);
                var retCol3 = GetDefaultAccent(ChartStyleId);
                var retCol4 = GetSchemeColorTint(eSchemeColor.Background1, 0.95d);

                return GetStyleColorOrDefault(ChartStyleId, retCol, retCol2, retCol3.Value, retCol4);
            }
            else
            {
                return null;
            }
        }

        protected Color GetEffectForChartElement(ChartElement element, int ChartStyleId)
        {
            throw new NotImplementedException("This method has not been implmented yet");
        }

    }
}
