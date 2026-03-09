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
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;
using OfficeOpenXml.Utils.String;
using OfficeOpenXml.Utils.TypeConversion;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Export.ImageRenderer.RenderItems.SvgItem
{
    internal class SvgChartSerieDataLabel : SvgChartObject
    {
        List<SvgTextBox> dlblTextBoxes = new List<SvgTextBox>();

        public SvgChartSerieDataLabel(SvgChart chart, ExcelChartSerieDataLabel dlblSerie, BoundingBox maxBounds) : base(chart)
        {
            foreach(var dlbl in dlblSerie.DataLabels)
            {
                var txtBox = new SvgTextBox(chart, maxBounds, maxBounds);
                txtBox.ImportTextBody(dlbl.TextBody);
                dlblTextBoxes.Add(txtBox);
            }
        }

        internal override void AppendRenderItems(List<RenderItem> renderItems)
        {
            SvgGroupItem groupItem = new SvgGroupItem(DrawingRenderer, Bounds);
            renderItems.Add(groupItem);

            foreach (var renderItem in dlblTextBoxes)
            {
                renderItem.AppendRenderItems(renderItems);
            }

            renderItems.Add(new SvgEndGroupItem(DrawingRenderer, Bounds));
        }

        //public override RenderItemType Type => throw new NotImplementedException();

        //public override void Render(StringBuilder sb)
        //{
        //    throw new NotImplementedException();
        //}
    }
}
