using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgChartDataLabelStandard : SvgChartObject
    {
        internal bool HasLegendKey { get; private set; } = false;

        bool _hasManualLayout = false;

        bool haveAdjustedForIcon = false;

        internal SvgTextBox TxtBox;

        public SvgChartDataLabelStandard(DrawingChart chart, string dataLabelText) : base(chart)
        {
            var txtBox = new SvgTextBox(chart, chart.Bounds, chart.Bounds);
            txtBox.AddText(0, dataLabelText);
            FitToTextBoxContent();
        }

        public SvgChartDataLabelStandard(DrawingChart chart, ExcelChartDataLabelStandard standard) : base(chart)
        {
            HasLegendKey = standard.ShowLegendKey;
        }

        public SvgChartDataLabelStandard(DrawingChart chart, ExcelChartDataLabelStandard standard, SvgTextBox txtBox) : base(chart)
        {
            HasLegendKey = standard.ShowLegendKey;
            TxtBox = txtBox;
            FitToTextBoxContent();
        }

        private void FitToTextBoxContent()
        {
            Bounds.Left = TxtBox.Left;
            Bounds.Top = TxtBox.Top;
            Bounds.Height = TxtBox.Height;
            Bounds.Width = TxtBox.Width;
        }

        internal void AddSeriesIcon(double iconWidth, double iconHeight)
        {
            if (haveAdjustedForIcon == false)
            {
                if (_hasManualLayout == false)
                {
                    TxtBox.Left += iconWidth + TxtBox.LeftMargin;
                    FitToTextBoxContent();
                    if (iconHeight > TxtBox.Height)
                    {
                        Bounds.Height = iconHeight;
                    }
                }
                else
                {
                    Bounds.Left += iconWidth + TxtBox.LeftMargin;
                    Bounds.Width += iconWidth;
                    Bounds.Height += iconHeight;
                }
                haveAdjustedForIcon = false;
            }
        }

        internal void ImportDataLabel(SvgChart chart, ExcelChartStandardSerie serie, ExcelChartDataLabelStandard dataLabel, object xValue, object yValue, ExcelDrawingParagraph defaultParagraph, BoundingBox maxBounds)
        {
            List<string> dlblStrings = new List<string>();

            if (dataLabel.ShowSeriesName)
            {
                dlblStrings.Add(serie.GetHeaderString());
            }
            if (dataLabel.ShowCategory)
            {
                dlblStrings.Add(xValue.ToString());
            }
            if (dataLabel.ShowValue)
            {
                dlblStrings.Add(yValue.ToString());
            }

            var separator = string.IsNullOrEmpty(dataLabel.Separator) ? "," : dataLabel.Separator;

            string finalString = "";
            for (int j = 0; j < dlblStrings.Count; j++)
            {
                finalString += dlblStrings[j];
                if (j != dlblStrings.Count - 1)
                {
                    finalString += separator;
                }
            }

            var txtBox = new SvgTextBox(chart, Bounds, maxBounds);
            txtBox.ImportTextBody(dataLabel.TextBody);

            if (txtBox.TextBody.Paragraphs.Count == 0)
            {
                txtBox.TextBody.ImportParagraph(defaultParagraph, 0, finalString);
                //txtBox.TextBody.AddParagraph(0, finalString);
            }
            else if (txtBox.TextBody.Paragraphs.Count == 1)
            {
                txtBox.TextBody.ImportParagraph(dataLabel.TextBody.Paragraphs[0], 0, finalString);
                //Remove dummy paragraph added by ImportTextBody
                txtBox.TextBody.Paragraphs.RemoveAt(0);
            }
            //Reset run y-position.
            //Datalabel does not use the standard line-spacing textbody offsets
            txtBox.TextBody.Paragraphs[0].Runs[0].YPosition = 0;

            TxtBox = txtBox;

            if (dataLabel is ExcelChartDataLabelItem)
            {
                var individualLabel = dataLabel as ExcelChartDataLabelItem;

                if (individualLabel.Layout != null && individualLabel.Layout.HasLayout)
                {
                    FitToTextBoxContent();
                    _hasManualLayout = true;
                    var rect = GetRectFromManualLayout(chart, individualLabel.Layout);
                    Rectangle = rect;
                    Bounds.Left = Rectangle.Left;
                    Bounds.Top = Rectangle.Top;
                    //Bounds.Width = Rectangle.Width;
                    //Bounds.Height = Rectangle.Height;
                }
            }
            else
            {
                FitToTextBoxContent();
            }
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            var group = new SvgGroupItem(ChartRenderer, Bounds);
            renderItems.Add(group);

            TxtBox.AppendRenderItems(renderItems);

            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }
    }
}
