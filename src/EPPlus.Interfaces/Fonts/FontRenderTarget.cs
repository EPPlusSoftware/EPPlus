/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2026         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// The kind of output text is being laid out for. Decides whether web font substitution applies.
    /// </summary>
    public enum FontRenderTarget
    {
        /// <summary>
        /// Output where EPPlus controls the fonts, e.g. PDF with embedded fonts. No substitution.
        /// </summary>
        Document,
        /// <summary>
        /// Output rendered by a browser using its own fonts, e.g. SVG and HTML.
        /// Fonts unlikely to be available to a browser are substituted.
        /// </summary>
        Web
    }
}