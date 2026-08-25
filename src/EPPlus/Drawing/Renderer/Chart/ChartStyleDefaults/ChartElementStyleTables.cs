using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.Encryption;
using System;
using System.Drawing;
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

    internal abstract class ChartDrawingObjectWithDefaults : ChartDrawingObject
    {
        public ChartDrawingObjectWithDefaults(ChartRenderer chart) : base(chart)
        {

        }

        private Color GetSchemeColorTint(eSchemeColor sColor, double tint = 0.0d)
        {
            if(tint < 0)
            {
                tint = 1 + tint;
            }
            else if(tint > 0)
            {
                tint = 1 - tint;
            }
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
                //Set to default instead for export
                //Otherwise epplus generated get weird.
                styleId = 2;
                //return Color.Empty;
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
        protected ExcelThemeLine GetThemedLine(ChartElement element, int ChartStyleId, bool nodeIsEmpty , out Color? lineColor)
        {
            //Chart style can only be above 48 if it is Style102 which in this case should be equivalent with style2
            //Alternatively it's an unkown or unset style which should also default to style2
            var styleId = ChartStyleId > (int)eChartStyle.Style48 ? (int)eChartStyle.Style2 : ChartStyleId;

            var AreaOrFloor = (ChartElement.ChartArea | ChartElement.Floor);
            if (AreaOrFloor.HasFlag(element))
            {
                var themedLine = ChartRenderer.Theme.FormatScheme.BorderStyle[0];

                //When the node exists but is empty Excel does not apply default styles
                //It directly applies the themedLineColor
                if (nodeIsEmpty)
                {
                    //Is node empty inside the theme
                    if(themedLine.HasFill == false)
                    {
                        lineColor = Color.Transparent;
                        return themedLine;
                    }

                    bool isSchemeColor = themedLine.Fill.SolidFill.Color.ColorType == eDrawingColorType.Scheme && themedLine.Fill.SolidFill.Color.SchemeColor.Color == eSchemeColor.Style;

                    if (isSchemeColor)
                    {
                        lineColor = GetDefaultBorderColorForElement(element, styleId);

                        if (themedLine.Fill.SolidFill.Color.Transforms.Count > 0 && lineColor.HasValue)
                        {
                            //var schemeClr = tc.ColorConverter.GetSchemeColor(ChartRenderer.Theme, eSchemeColor.Dark1);
                            //var tint = GetSchemeColorTint(eSchemeColor.Dark1, 0.45d);
                            lineColor = tc.ColorConverter.ApplyTransforms(lineColor.Value, themedLine.Fill.SolidFill.Color.Transforms);
                        }

                        return themedLine;
                    }
                    else
                    {
                        lineColor = themedLine.Fill.Color;
                    }
                    return themedLine;
                }

                lineColor = GetDefaultBorderColorForElement(element, styleId);

                var themeColor = tc.ColorConverter.GetThemeColor(ChartRenderer.Theme, themedLine.Fill.SolidFill.Color);
                if (themedLine.HasFill == false)
                {
                    //Node exists but has no fill. Excel considers this the same as transparent/noFill
                    lineColor = Color.Transparent;
                    return themedLine;
                }

                if (styleId < 41)
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
            //Chart style can only be above 48 if it is Style102 which in this case should be equivalent with style2
            //Alternatively it's an unkown or unset style which should also default to style2
            var styleId = ChartStyleId > (int)eChartStyle.Style48 ? (int)eChartStyle.Style2 : ChartStyleId;

            var bg = ChartRenderer.Theme.FormatScheme.BackgroundFillStyle[0];

            if ((ChartElement.Floor | ChartElement.Walls).HasFlag(element))
            {
                fillColor = GetDefaultFillColorForElement(element, styleId);
                var themedFill = ChartRenderer.Theme.FormatScheme.BackgroundFillStyle[0];

                if (styleId > 32)
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
                    $"Only Walls or Floor has a default themed fill");
            }
        }

        protected Color? GetDefaultBorderColorForElement(ChartElement element, int ChartStyleId)
        {
            //return Color.Empty;

            if((ChartElement.Axis | ChartElement.MajorGridLines).HasFlag(element))
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
            else if ((ChartElement.ChartArea | ChartElement.DataTable | ChartElement.Floor).HasFlag(element))
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
            else if((ChartElement.Floor | ChartElement.Walls | ChartElement.PlotArea2d).HasFlag(element))
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


        abstract internal Color? GetDefaultFillColor();
        abstract internal Color? GetDefaultBorderColor();
    }
}
