/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  23/11/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.DrawingRenderer.Svg
{
    /// <summary>
    /// Options for rendering drawings to svg.
    /// </summary>
    public class SvgRenderOptions
    {
        /// <summary>
        /// The width of the drawing in pixels used for calculating output image. Overrides the width of the drawing if set. If not set, the width will be calculated based on the drawing.
        /// For output sizing, use the <see cref="SvgSize"/> property instead.
        /// </summary>
        public int? Width { get; set; }
        /// <summary>
        /// The height of the drawing in pixels. Overrides the height of the drawing if set. If not set, the width will be calculated based on the drawing.
        /// For output sizing, use the <see cref="SvgSize"/> property instead.
        /// </summary>
        public int? Height { get; set; }
        /// <summary>
        /// Sets the width and height of the svg image. If not set, the size will be calculated based on the drawings dimensions.
        /// </summary>
        public SvgSize SvgSize { get; } = new SvgSize();

        /// <summary>
        /// The render target used for text layout. Only has an effect when <see cref="EmbedFonts"/> is
        /// false. <see cref="FontRenderTarget.Web"/> (the default) then substitutes fonts unlikely to be
        /// available to a browser, both when measuring and in the output, e.g. Aptos Narrow with Calibri.
        /// <see cref="FontRenderTarget.Document"/> keeps the requested fonts, for when they are made
        /// available to the consumer by other means.
        /// </summary>
        public FontRenderTarget FontTarget { get; set; } = FontRenderTarget.Web;

        /// <summary>
        /// When true (the default), text is measured with the requested fonts (no web font
        /// substitution) and the fonts are embedded in the svg as subsets that only hold the glyphs
        /// used, so the output does not depend on the fonts the browser has. A font is only embedded
        /// when its licence allows it, see <c>IEpplusFontConfiguration.OnFontEmbedding</c>; other fonts
        /// are referenced by name. When false, no font is embedded and <see cref="FontTarget"/>
        /// decides whether fonts are substituted.
        /// </summary>
        public bool EmbedFonts { get; set; } = true;
    }
    
}