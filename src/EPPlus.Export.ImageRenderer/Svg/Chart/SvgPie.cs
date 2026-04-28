using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Chart
{
    internal class SvgPie : SvgRenderItem
    {
        double _pieExplosionPercent;


        public SvgPie(DrawingBase renderer, BoundingBox parent, ExcelPieChart chart, SvgGroupItemNew container) : base(renderer, parent)
        {

        }

        public override RenderItemType Type => RenderItemType.Group;

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            throw new NotImplementedException();
        }
    }
}
