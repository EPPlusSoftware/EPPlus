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
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Style.Effect;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
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
        internal string TransformOffset = "";

        public override void Render(StringBuilder sb)
        {
            int width = (int)Bounds.Width;
            int height = (int)Bounds.Height;

            sb.Append($"<path d=\"");
            for (int i = 0; i < Commands.Count; i++)
            {
                Commands[i].Render(width, height, sb);
            }
            sb.Append("\" ");
            base.Render(sb);
            sb.Append("/>");
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