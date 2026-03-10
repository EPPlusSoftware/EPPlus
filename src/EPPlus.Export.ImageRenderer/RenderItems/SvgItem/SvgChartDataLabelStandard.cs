using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgChartDataLabelStandard : SvgChartObject
    {
        internal bool hasLegendKey { get; private set; } = false;

        internal SvgTextBox TxtBox;

        public SvgChartDataLabelStandard(DrawingChart chart, ExcelChartDataLabelStandard standard, SvgTextBox txtBox) : base(chart)
        {
            hasLegendKey = standard.ShowLegendKey;
            TxtBox = txtBox;
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            throw new NotImplementedException();
        }
    }
}
