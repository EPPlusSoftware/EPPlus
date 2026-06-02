using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    //internal class SvgChartLegendIcon : SvgRenderLineItem
    //{
    //    const float MarginExtra = 1.5f;
    //    const float MiddleMargin = 7.5f;
    //    const float LineLength = 21;

    //    public SvgChartLegendIcon(DrawingChart sc, SvgRenderRectItem Rectangle, ExcelChartStandardSerie s, TextMeasurement tm, double topMargin, double leftMargin, TextMeasurement sHM, SvgLegendSerie pSls) : base(sc, Rectangle.Bounds)
    //    {
    //        SetDrawingPropertiesFill(s.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
    //        SetDrawingPropertiesBorder(s.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, s.Border.Fill.Style != eFillStyle.NoFill, 0.75);

    //        if (sc.Chart.Legend.Position == eLegendPosition.Top ||
    //           sc.Chart.Legend.Position == eLegendPosition.Bottom)
    //        {
    //            float y = (float)Rectangle.Top + (float)topMargin + tm.Height / 2 + MarginExtra;
    //            float x = 0;
    //            if (pSls == null)
    //            {
    //                x = (float)Rectangle.Left + (float)leftMargin;// + MarginIconText;
    //            }
    //            else
    //            {
    //                x = (float)pSls.Textbox.Bounds.Right + MiddleMargin;
    //            }

    //            X1 = x;
    //            Y1 = y;
    //            X2 = x + LineLength;
    //            Y2 = y;
    //            LineCap = eLineCap.Round;
    //        }
    //        else
    //        {
    //            double y;
    //            if (pSls == null)
    //            {
    //                y = topMargin + tm.Height / 2 + MarginExtra;
    //            }
    //            else
    //            {
    //                var pTm = sHM;
    //                y = ((SvgRenderLineItem)pSls.SeriesIcon).Y1 + pTm.Height / 2 + tm.Height / 2 + MiddleMargin;
    //            }

    //            X1 = (float)leftMargin; //4
    //            Y1 = y;
    //            X2 = (float)LineLength;
    //            Y2 = y;
    //            LineCap = eLineCap.Round;
    //        }
    //    }

    //    //internal override void AppendRenderItems(List<RenderItem> renderItems)
    //    //{
    //    //    throw new NotImplementedException();
    //    //}

    //    //private SvgRenderLineItem GetLineSeriesIcon(DrawingChart sc, ExcelChartStandardSerie s, TextMeasurement sHM, TextMeasurement tm, SvgLegendSerie pSls)
    //    //{
    //    //    var item = new SvgRenderLineItem(sc, Rectangle.Bounds);
    //    //    item.SetDrawingPropertiesFill(s.Fill, sc.Chart.StyleManager.Style.SeriesLine.FillReference.Color);
    //    //    item.SetDrawingPropertiesBorder(s.Border, sc.Chart.StyleManager.Style.SeriesLine.BorderReference.Color, s.Border.Fill.Style != eFillStyle.NoFill, 0.75);

    //    //    if (sc.Chart.Legend.Position == eLegendPosition.Top ||
    //    //       sc.Chart.Legend.Position == eLegendPosition.Bottom)
    //    //    {
    //    //        float y = (float)Rectangle.Top + (float)TopMargin + tm.Height / 2 + MarginIconText;
    //    //        float x = 0;
    //    //        if (pSls == null)
    //    //        {
    //    //            x = (float)Rectangle.Left + (float)LeftMargin;// + MarginIconText;
    //    //        }
    //    //        else
    //    //        {
    //    //            x = (float)pSls.Textbox.Bounds.Right + MarginHeight;
    //    //        }

    //    //        item.X1 = x;
    //    //        item.Y1 = y;
    //    //        item.X2 = x + LineLength;
    //    //        item.Y2 = y;
    //    //        item.LineCap = eLineCap.Round;
    //    //    }
    //    //    else
    //    //    {
    //    //        double y;
    //    //        if (pSls == null)
    //    //        {
    //    //            y = TopMargin + tm.Height / 2 + MarginIconText;
    //    //        }
    //    //        else
    //    //        {
    //    //            var pTm = sHM;
    //    //            y = ((SvgRenderLineItem)pSls.SeriesIcon).Y1 + pTm.Height / 2 + tm.Height / 2 + MarginHeight;
    //    //        }

    //    //        item.X1 = (float)LeftMargin; //4
    //    //        item.Y1 = y;
    //    //        item.X2 = (float)LineLength;
    //    //        item.Y2 = y;
    //    //        item.LineCap = eLineCap.Round;
    //    //    }

    //    //    return item;
    //    //}
    //}
}
