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
using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Style.Coloring;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.TypeConversion;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using tc = OfficeOpenXml.Utils.TypeConversion;

namespace EPPlusImageRenderer.Svg
{
    internal abstract class ChartDrawingObject : DrawingObject
    {
        internal ChartRenderer ChartRenderer;
        internal ExcelChart Chart => (ExcelChart)ChartRenderer.Drawing;

        internal RenderContext RenderContext => ChartRenderer.RenderContext;
        internal ChartDrawingObject(ChartRenderer chart)
        {
            ChartRenderer = chart;
            //Fixes null ref but might be inaccurate for some objects...
            Rectangle = new RectRenderItem(chart.Bounds);
            //InitStyleColors();
        }
        internal void SetMargins(ExcelTextBody tb)
        {
            tb.GetInsetsOrDefaults(out double l, out double r, out double t, out double b);
            LeftMargin = l;
            RightMargin = r;
            TopMargin = t;
            BottomMargin = b;
        }
        internal double LeftMargin { get; set; }
        internal double RightMargin { get; set; }
        internal double TopMargin { get; set; }
        internal double BottomMargin { get; set; }
        internal virtual RectRenderItem Rectangle { get; set; }
        internal virtual Color? DefaultFillColor { get; }
        internal virtual Color? DefaultBorderColor { get; }
        protected static RectRenderItem GetRectFromManualLayout(ChartRenderer sc, ExcelLayout layout, BoundingBox parent=null)
        {
            var bounds = parent ?? sc.Bounds;
            var rect = new RectRenderItem(sc.ChartArea.Rectangle.Bounds);
            var ml = layout.ManualLayout;
            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.Left = bounds.Width * (float)(layout.ManualLayout.Left ?? 0D) / 100;
            }
            else
            {
                rect.Left = bounds.Width * (float)(ml.Left ?? 0D) / 100;
                //TODO:Add factor from default position
            }

            //Width is always factor.
            rect.Width = bounds.Width * ml.GetWidth() / 100;

            if (ml.LeftMode == eLayoutMode.Edge)
            {
                rect.Top = bounds.Height * (float)(layout.ManualLayout.Top ?? 0D) / 100;
            }
            else
            {
                rect.Top = bounds.Height * (float)(ml.Top ?? 0D) / 100;
                //TODO:Add factor from default position
            }
            //Height is always factor.
            rect.Height = bounds.Height * ml.GetHeight() / 100;
            return rect;
        }
        /// <summary>
        /// Get the X serie values. If the values are not numeric, return a serie with the index values (1,2,3,...). Trendline calculation requires numeric X values, but Excel allows non-numeric X values for trendlines, in which case it uses the index values as X for calculation.
        /// </summary>
        /// <param name="xSerie">Input values</param>
        /// <returns>Output doubles</returns>
        internal List<double> GetXSerie(List<object> xSerie)
        {
            var l = new List<double>();
            for (int i = 0; i < xSerie.Count; i++)
            {
                if (ConvertUtil.IsExcelNumeric(xSerie[i]))
                {
                    l.Add(ConvertUtil.GetValueDouble(xSerie[i]));
                }
                else
                {
                    return xSerie.Select((x, index) => (double)(index + 1)).ToList();
                }
            }
            return l;
        }

        /// <summary>
        /// Default style color for style 1-32
        /// </summary>
        protected internal Color StyleColor1 { get; internal set; }
        /// <summary>
        /// Styles 33-34
        /// </summary>
        protected internal Color StyleColor2 { get; internal set; }
        /// <summary>
        /// Styles 35-40
        /// </summary>
        protected internal Color StyleColor3 { get; internal set; }
        /// <summary>
        /// Styles 41-48
        /// </summary>
        protected internal Color StyleColor4 { get; internal set; }


        /// <summary>
        /// Default style color for style 1-32
        /// </summary>
        protected internal Color StyleBorderColor1 { get; internal set; }
        /// <summary>
        /// Styles 33-34
        /// </summary>
        protected internal Color StyleBorderColor2 { get; internal set; }
        /// <summary>
        /// Styles 35-40
        /// </summary>
        protected internal Color StyleBorderColor3 { get; internal set; }
        /// <summary>
        /// Styles 41-48
        /// </summary>
        protected internal Color StyleBorderColor4 { get; internal set; }

        protected Color GetThemeColorTint(eThemeSchemeColor themeColor, double tint = 0.0d )
        { 
            var schemeClr = tc.ColorConverter.GetThemeColor(ChartRenderer.Theme, themeColor);
            var tintedSchemeColor = tc.ColorConverter.ApplyTintDrawing(schemeClr, tint);
            return tintedSchemeColor;
        }

        //internal abstract void InitStyleColors();

        /// <summary>
        /// This function provides the default chart color for a given chart object
        /// </summary>
        /// <param name="styleId">Chart style Id</param>
        /// <returns></returns>
        internal Color? GetStyleColorOrDefault(int styleId)
        {
            Color? themeColor = null;
            styleId = styleId > (int)eChartStyle.Style48 ? (int)eChartStyle.Style2 : styleId;

            if (styleId == 0)
            {
                return Color.Empty;
            }

            var bg = ChartRenderer.Theme.FormatScheme.BackgroundFillStyle[0];

            themeColor = tc.ColorConverter.GetThemeColor(ChartRenderer.Theme, eThemeSchemeColor.Background1);

            if (bg.SolidFill.Color.ColorType == eDrawingColorType.Scheme && bg.SolidFill.Color.SchemeColor.Color == eSchemeColor.Style)
            {
                if(styleId <= 32)
                {
                    themeColor = StyleColor1;
                }
                else if(styleId <= 34)
                {
                    themeColor = StyleColor2;
                }
                else if(styleId <= 40)
                {
                    themeColor = StyleColor3;
                }
                else if(styleId <= 48)
                {
                    themeColor = StyleColor4;
                }

                //if (styleId <= 40)
                //{
                //    //Text1 AKA dk1 (in standard case)
                //    themeColor = tc.ColorConverter.GetThemeColor(ChartRenderer.Theme, eThemeSchemeColor.Text1);

                //    //var bg1Col = tc.ColorConverter.GetThemeColor(Theme, eThemeSchemeColor.Background1);

                //    if (bg.SolidFill.Color.Transforms.Count > 0)
                //    {
                //        //themeColor = tc.ColorConverter.ApplyTintDrawing(themeColor.Value, 0.15d);
                //        //but even in this case if there is no ln node found in style it appears to default to 75% despite a scheme color existing in the theme
                //        themeColor = tc.ColorConverter.ApplyTransforms(themeColor.Value, bg.SolidFill.Color.Transforms);
                //    }
                //    else
                //    {
                //        //Default value Should arguably be 75% tint themeColor but something is strange...
                //        //It appears closer to 50% in this specific case
                //        //It also appears to be tx1 (black) and apply color and tint 0.25 in vba
                //        var newTheme = tc.ColorConverter.ApplyTintDrawing(themeColor.Value, 0.25d);
                //        themeColor = newTheme;

                //    }
                //}
                //else
                //{
                //    //41-48
                //    //aka light1
                //    themeColor = tc.ColorConverter.GetThemeColor(ChartRenderer.Theme, eThemeSchemeColor.Background1);
                //    //themedLine = null;
                //}
            }
            return themeColor;
        }
    }
}
