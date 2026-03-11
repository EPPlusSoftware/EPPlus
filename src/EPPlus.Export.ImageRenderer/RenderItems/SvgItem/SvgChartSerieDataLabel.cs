using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg.Chart.Util;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgChartSerieDataLabel : DrawingObjectNoBounds
    {
        //positioning is handled by parent item via these
        internal List<SvgGroupItem> groupItems = new List<SvgGroupItem>();

        private RenderItem seriesIcon = null;
        private List<SvgChartDataLabelStandard> dataLabels = new List<SvgChartDataLabelStandard>();
        
        string separator;

        ExcelTextFont defaultFont;
        ExcelDrawingParagraph defaultParagraph;

        public SvgChartSerieDataLabel(SvgChart chart, ExcelChartSerieDataLabel dlblSerie, BoundingBox maxBounds, ExcelChartStandardSerie serie, List<object> xValues, List<object> yValues, int index) : base(chart)
        {
            bool addSeriesIcon = false;

            if(dlblSerie.TextBody.Paragraphs.Count != 0)
            {
                defaultParagraph = dlblSerie.TextBody.Paragraphs[0];
                defaultFont = dlblSerie.TextBody.Paragraphs[0].DefaultRunProperties;
            }

            if (dlblSerie.DataLabels.Count == 0 && serie.NumberOfItems > 0)
            {

                for (int i = 0; i < serie.NumberOfItems; i++)
                {
                    AddDatalabel(chart, serie, dlblSerie, xValues[i], yValues[i], maxBounds, ref addSeriesIcon);
                }
            }
            else
            {
                for (int i = 0; i < dlblSerie.DataLabels.Count; i++)
                {
                    var dataLabel = dlblSerie.DataLabels[i];

                    AddDatalabel(chart, serie, dataLabel, xValues[i], yValues[i], maxBounds, ref addSeriesIcon);
                }
            }

            if(addSeriesIcon)
            {
                if (chart.Legend == null)
                {
                    seriesIcon = chart.GetSeriesIcon(serie, index, maxBounds);
                }
                else
                {
                    var legendItem = chart.Legend;
                    var seriesIconOrig = (SvgRenderLineItem)legendItem.SeriesIcon[index].SeriesIcon;
                    var clonedIcon = seriesIconOrig.Clone(chart);

                    clonedIcon.Y1 = 0;
                    clonedIcon.Y2 = 0;

                    seriesIcon = clonedIcon;
                }
            }

        }

        private void AddDatalabel(SvgChart chart, ExcelChartStandardSerie serie, ExcelChartDataLabelStandard dataLabel, object xValue, object yValue, BoundingBox maxBounds, ref bool addSeriesIcon)
        {
            if (addSeriesIcon == false && dataLabel.ShowLegendKey)
            {
                addSeriesIcon = dataLabel.ShowLegendKey;
            }

            var newDataLabel = new SvgChartDataLabelStandard(chart, dataLabel);
            newDataLabel.ImportDataLabel(chart, serie, dataLabel, xValue, yValue, defaultParagraph, maxBounds);
            dataLabels.Add(newDataLabel);
        }

        internal void SetPositionOffset(double xPos, double yPos, int i)
        {
            dataLabels[i].SetOriginPointOffset(xPos, yPos);
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            for(int i = 0; i< dataLabels.Count; i++) 
            {
                if (seriesIcon != null && dataLabels[i].HasLegendKey)
                {
                    //groupItems[i].Bounds.Left += (seriesIcon.Bounds.Width / 2);
                    //groupItems[i].GroupTransform = $"transform=\"translate({groupItems[i].Bounds.Left.PointToPixelString()}, {groupItems[i].Bounds.Top.PointToPixelString()})\"";
                    dataLabels[i].AddSeriesIcon(seriesIcon.Bounds.Width, seriesIcon.Bounds.Height);
                    //renderItems.Add(groupItems[i]);
                    renderItems.Add(seriesIcon);
                }
                else
                {
                    //renderItems.Add(groupItems[i]);
                }

                dataLabels[i].AppendRenderItems(renderItems);

                //renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
            }
        }
    }
}
