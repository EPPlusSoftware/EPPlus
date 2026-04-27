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
using EPPlus.Export.ImageRenderer.Utils;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgRenderPathItem : SvgRenderItem
    {
        public SvgRenderPathItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {
            Bounds.Width = parent.Width;
            Bounds.Height = parent.Height;
        }
        public override RenderItemType Type { get => RenderItemType.Path; }
        public List<PathCommands> Commands { get; set; } = new List<PathCommands>();

        public override void Render(StringBuilder sb)
        {
            //Draw transparent lines to create the compond line effect, as SVG does not support compound lines natively
            switch (CompoundLineStyle)
            {
                case eCompoundLineStyle.Single:
                    RenderPathItem(sb, null, null, null);
                    break;
                case eCompoundLineStyle.Double:
                    var name = $"double-stroke-{Guid.NewGuid().ToString()}";
                    sb.Append($"<defs><mask id=\"{name}\">");

                    RenderPathItem(sb, BorderWidth, "white", null);
                    RenderPathItem(sb, BorderWidth * (3D / 7D), "black", null);
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{BorderColor}\" mask=\"url(#{name})\" />");
                    break;
                case eCompoundLineStyle.DoubleThickThin:
                    WriteThickThin(sb, (BorderWidth??1D) * 1D / 7D);
                    break;
                case eCompoundLineStyle.DoubleThinThick:
                    WriteThickThin(sb, ((BorderWidth ?? 1D) * 1D / 7D) * -1);
                    break;
                case eCompoundLineStyle.TripleThinThickThin:
                    var guid = Guid.NewGuid().ToString();
                    var gapOffset = 5 * BorderWidth.Value / 16;
                    name = $"triple-stroke-{guid}";
                    sb.Append($"<defs>");
                    sb.Append($"<filter id=\"gap-left-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\"><feOffset dx=\"0\" dy=\"-{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>");
                    sb.Append($"<filter id=\"gap-right-{guid}\" x=\"-500%\" y=\"-500%\" width=\"1100%\" height=\"1100%\"><feOffset dx=\"0\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\" /></filter>");
                    sb.Append($"<mask id=\"{name}\">");
                    RenderPathItem(sb, BorderWidth, "white", null);
                    RenderPathItem(sb, BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-left-{guid})\"");
                    RenderPathItem(sb, BorderWidth * (1D / 8D), "black", $"filter=\"url(#gap-right-{guid})\"");
                    sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{BorderColor}\" mask=\"url(#{name})\" />");
                    break;
            }
        }

        private void WriteThickThin(StringBuilder sb, double gapOffset)
        {
            var guid = Guid.NewGuid().ToString();
            var name = $"double-thick-thin-stroke-{guid}";
            string gapFilterName = $"f-gap-shift-{guid}";
            sb.Append("<defs>");
            sb.Append($"<filter id=\"{gapFilterName}\" x=\"-50%\" y=\"-50%\" width=\"200%\" height=\"200%\"><feOffset in=\"SourceGraphic\" dy=\"{gapOffset.PointToPixel().ToString(CultureInfo.InvariantCulture)}\"/></filter>");
            sb.Append($"<mask id=\"{name}\">");
            RenderPathItem(sb, BorderWidth, "white", null);
            RenderPathItem(sb, BorderWidth * (1 / 4D), "black", $"filter=\"url(#{gapFilterName})\"");
            sb.Append($"</mask></defs><rect width=\"100%\" height=\"100%\" fill=\"{BorderColor}\" mask=\"url(#{name})\" />");
        }

        private void RenderPathItem(StringBuilder sb, double? borderWidth, string color, string filter)
        {
            var width = Bounds.Width.PointToPixel();
            var height = Bounds.Height.PointToPixel();

            sb.Append($"<path d=\"");
            for (int i = 0; i < Commands.Count; i++)
            {
                Commands[i].Render(width, height, sb);
            }
            sb.Append("\" ");
            RenderCompoundItems(sb, borderWidth, color, filter);

        }

        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            var clone = new SvgRenderPathItem(svgDocument, svgDocument.Bounds);
            CloneBase(clone);
            clone.Commands = CloneCommands(Commands);
            return clone;
        }

        private List<PathCommands> CloneCommands(List<PathCommands> commands)
        {
            var cloneList = new List<PathCommands>();
            foreach (var cmd in commands)
            {
                cloneList.Add(cmd.Clone());
            }
            return cloneList;
        }
        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = it = ir = ib = 0;
            foreach(var c in Commands)
            {
                if(c.Type==PathCommandType.Arc)
                {
                    var wR = c.Coordinates[0];
                    var hR = c.Coordinates[1];
                    var x = c.Coordinates[5];
                    var y = c.Coordinates[6];
                    if (x < il) il = x + wR;    //TODO: Maybe adjust with start- and swing- angle
                    if (x > ir) ir = x + wR;
                    if (y < it) it = y + hR;
                    if (y > ib) ib = y + hR;
                }
                else
                {
                    for(int i = 0; i < c.Coordinates.Length;i++)
                    {
                        if (i % 2 == 0)
                        {
                            if (c.Coordinates[i] < il) il = c.Coordinates[i];
                            if (c.Coordinates[i] > ir) ir = c.Coordinates[i];
                        }
                        else
                        {
                            if (c.Coordinates[i] < it) it = c.Coordinates[i];
                            if (c.Coordinates[i] > ib) ib = c.Coordinates[i];
                        }
                    }
                }
            }
        }
    }
}