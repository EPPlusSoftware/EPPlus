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
using System.Drawing;

namespace EPPlus.DrawingRenderer.RenderItems
{
    public class RenderShadowEffect : RenderStyle
    {
        public Color OuterShadowEffectColor { get; set; }
        public double Distance { get; set; }
        public double? BlurRadius { get; set; }
        public double? Direction { get; set; }

        public override string GetKey()
        {
            return $"{OuterShadowEffectColor.ToArgb()} {Distance} {BlurRadius} {Direction}";
        }
    }
}