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
using EPPlusImageRenderer.Svg;
using System.Text;

namespace EPPlusImageRenderer.RenderItems
{
    internal class SvgGroupItem : SvgRenderItem
    {
        public override SvgItemType Type => SvgItemType.Group;

        public string GroupTransform = "";

        internal SvgGroupItem() : base()
        {

        }
        internal SvgGroupItem(double rotation, double cx, double cy) : base()
        {
            if(rotation!=0)
            {
                if (cx == 0 && cy == 0)
                {
                    GroupTransform = $"transform(rotation({rotation}))";
                }
                else
                {
                    GroupTransform = $"transform(rotation({rotation}, {cx}, {cy})";
                }
            }
        }

        public override void Render(StringBuilder sb)
        {
            sb.Append($"<g {GroupTransform}>");
        }
        internal void RenderEndGroup(StringBuilder sb)
        {
            sb.Append($"</g>");
        }


        internal override SvgRenderItem Clone(SvgShape svgDocument)
        {
            return this.Clone(svgDocument);
        }

        internal override void GetBounds(out double il, out double it, out double ir, out double ib)
        {
            il = 0;
            it = 0;
            ir = 1;
            ib = 1;
        }
    }
}
