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
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlusImageRenderer.RenderItems;

namespace EPPlusImageRenderer.Svg
{
    internal class SvgLegendSerie
    {
        internal RenderItem SeriesIcon { get; set; }
        internal RenderItem MarkerIcon { get; set; }
        internal RenderItem MarkerBackground { get; set; }
        internal TextBodyItem Textbox { get; set;}
    }
}