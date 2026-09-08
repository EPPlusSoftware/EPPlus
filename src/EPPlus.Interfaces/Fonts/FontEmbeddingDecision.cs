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
    /// The action EPPlus takes for a font when preparing it for embedding.
    /// Returned from the callback registered via
    /// <see cref="IEpplusFontConfiguration.OnFontEmbedding"/>.
    /// </summary>
    public enum FontEmbeddingDecision
    {
        /// <summary>
        /// Follow the font's declared fsType: throw for a Restricted License font,
        /// embed whole for a no-subsetting font, subset otherwise.
        /// </summary>
        Default,
        /// <summary>
        /// Subset the font regardless of fsType. By choosing this, the caller asserts
        /// they hold the rights to embed and subset the font.
        /// </summary>
        Subset,
        /// <summary>Embed the whole font without subsetting.</summary>
        EmbedWhole,
        /// <summary>Do not embed the font; a fallback/substitute is used instead.</summary>
        Skip,
    }
}
