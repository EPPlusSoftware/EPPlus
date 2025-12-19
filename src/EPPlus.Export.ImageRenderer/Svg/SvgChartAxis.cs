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
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Collections.Generic;
using System.Text;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgChartAxis : SvgChartObject, IDrawingChartAxis
    {
        internal SvgChartAxis(SvgChart sc, ExcelChartAxisStandard ax) : base(sc.Chart)
        {
            if (sc.Chart.Series.Count == 0 || ax.Deleted==true)
            {
                return;
            }
            float textHeight, textWidth;

            SetMargins(ax.TextBody);

            if (ax.Layout.HasLayout)
            {
                Rectangle = GetRectFromManualLayout(sc, ax.Layout);
            }
            else
            {
                Rectangle = new SvgRenderRectItem(sc.Chart);
                //var w = 10; //Width/Height
                //switch(ax.AxisPosition)
                //{
                //    case eAxisPosition.Left:
                //        Rectangle.Y = 15;
                //        Rectangle.X = 0;
                //        break;
                //    case eAxisPosition.Right:
                //        Rectangle.Y = 0;
                //        Rectangle.X = 0;
                //        break;
                //    case eAxisPosition.Top:
                //        Rectangle.Y = 0;
                //        Rectangle.X = 0;
                //        break;
                //    case eAxisPosition.Bottom:
                //        Rectangle.Y = 0;
                //        Rectangle.X = 0;
                //        break;
                //}
                //Rectangle.Width = sc.Plotarea.Rectangle.Width;
                //Rectangle.Height = w;
                //Rectangle.Width = 0;
            }
            Values = GetAxisValue(ax, Rectangle);

            Rectangle.SetDrawingPropertiesFill(ax.Fill, sc.Chart.StyleManager.Style.Title.FillReference.Color);
            Rectangle.SetDrawingPropertiesBorder(ax.Border, sc.Chart.StyleManager.Style.Title.BorderReference.Color, ax.Border.Fill.Style!=eFillStyle.NoFill, 0.75);
        }

        public List<object> Values
        {
            get;
            private set;
        }
        public SvgChartTitle AxisTitle { get; }

        public override RenderItemType Type => throw new System.NotImplementedException();

        public override void Render(StringBuilder sb)
        {
            AxisTitle?.Render(sb);
            Rectangle.Render(sb);
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            throw new System.NotImplementedException();
        }
    }
}