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
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Utils.TypeConversion;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace EPPlusImageRenderer.Svg
{
    internal abstract class ChartDrawingObject : DrawingObject
    {
        internal ChartRenderer ChartRenderer;
        internal ExcelChart Chart => (ExcelChart)ChartRenderer.Drawing;
        internal ChartDrawingObject(ChartRenderer chart)
        {
            ChartRenderer = chart;
            //Fixes null ref but might be inaccurate for some objects...
            Rectangle = new RectRenderItem(chart.Bounds);
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


    }
}
