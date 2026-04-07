using EPPlusImageRenderer;
using EPPlusImageRenderer.RenderItems;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.DefinitionUtils
{
    enum LinePatternType
    {
        Vertical,
        Horizontal
    }

    internal class LinePattern : PatternItem
    {
        LinePatternType type;

        internal SvgRenderLineItem LineItem;

        public LinePattern(DrawingBase baseRend, string id, LinePatternType linesType) : base(baseRend, id)
        {
            type = linesType;
            LineItem = new SvgRenderLineItem(baseRend, Bounds);

            LineItem.BorderWidth = 2;
            LineItem.Suffix = "%";

            LineItem.BorderColor = "#" + Color.DarkGray.To6CharHexString();

            switch (type)
            {
                case LinePatternType.Vertical:
                    LineItem.X1 = 0; LineItem.X2 = 0; LineItem.Y1 = 0; LineItem.Y2 = 100;
                    break;
                case LinePatternType.Horizontal:
                    LineItem.X1 = 0; LineItem.X2 = 100; LineItem.Y1 = 0; LineItem.Y2 = 0;
                    break;
            }

            SetNumberOfLines(6);
        }

        public override RenderItemType Type => RenderItemType.Reference;

        /// <summary>
        /// Sets number of lines via the width or height percent
        /// </summary>
        internal void SetNumberOfLines(int numberOfLines)
        {
            switch (type)
            {
                case LinePatternType.Vertical:
                    widthPercent = (double)numberOfLines / 100d;
                    break;
                  case LinePatternType.Horizontal:
                    heightPercent = (double)numberOfLines / 100d;
                    break;
            }
        }

        public override void Render(StringBuilder sb)
        {
            _items.Add(LineItem);
            base.Render(sb);
        }
    }
}
