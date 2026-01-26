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
using EPPlus.Graphics;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml.Drawing.Theme;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Drawing;
using System.Globalization;
using System.Text;

namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgEndGroupItem : RenderItem
    {
        internal SvgEndGroupItem(DrawingBase renderer, BoundingBox parent) : base(renderer, parent)
        {

        }
        public override RenderItemType Type => throw new System.NotImplementedException();
        public override void Render(StringBuilder sb)
        {
                sb.Append("</g>");
        }
    }
    internal class SvgGroupItem : RenderItem
    {
        public override RenderItemType Type => RenderItemType.Group;

        public string GroupTransform = "";

        internal SvgGroupItem(DrawingBase renderer, BoundingBox bounds) : base(renderer)
        {
            Bounds = bounds;
            if (Bounds.Left != 0 || Bounds.Top != 0)
            {
                GroupTransform = $"transform=\"translate({Bounds.Left.ToString(CultureInfo.InvariantCulture)}, {Bounds.Top.ToString(CultureInfo.InvariantCulture)})\"";
            }
        }
        internal SvgGroupItem(DrawingBase renderer, BoundingBox parent, double rotation) : base(renderer, parent)
        {
            var bounds = "";
            var rot = "";
            if(Bounds.Left!=0 || Bounds.Top!=0)
            {
                bounds = $"translate({Bounds.Left.ToString(CultureInfo.InvariantCulture)}, {Bounds.Top.ToString(CultureInfo.InvariantCulture)})";
            }
            if(rotation!=0)
            {
                var cx = parent.Width / 2 + parent.Left;
                var cy = parent.Height / 2 + parent.Top;
                if (cx == 0 && cy == 0)
                {
                    rot = $"rotate({rotation.ToString(CultureInfo.InvariantCulture)}))";
                }
                else
                {
                    rot = $"rotate({rotation.ToString(CultureInfo.InvariantCulture)}, {cx.ToString(CultureInfo.InvariantCulture)}, {cy.ToString(CultureInfo.InvariantCulture)})";
                }
            }

            if(string.IsNullOrEmpty(bounds)==false && string.IsNullOrEmpty(rot)==false)
            {
                GroupTransform = $"transform=\"{bounds} {rot}\"";
            }
            else if(string.IsNullOrEmpty(bounds)==false)
            {
                GroupTransform = $"transform=\"{bounds}\"";
            }
            else if (string.IsNullOrEmpty(rot) == false)
            {
                GroupTransform = $"transform=\"{rot}\"";
            }
        }

        public override void Render(StringBuilder sb)
        {
            if(string.IsNullOrEmpty(GroupTransform))
            {
                sb.Append($"<g>");
            }
            else
            {
                sb.Append($"<g {GroupTransform}>");
            }
        }

        //internal override SvgRenderItem Clone(SvgShape svgDocument)
        //{
        //    return this.Clone(svgDocument);
        //}

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = 0;
            it = 0;
            ir = 1;
            ib = 1;
        }
    }
}
