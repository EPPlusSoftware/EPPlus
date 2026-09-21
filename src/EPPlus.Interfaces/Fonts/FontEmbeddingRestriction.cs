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
    /// The embedding/subsetting restriction a font declares via its OS/2 fsType field.
    /// This is a pure interpretation of fsType — it carries no policy about what EPPlus does.
    /// </summary>
    public enum FontEmbeddingRestriction
    {
        /// <summary>Font may be embedded and subsetted freely.</summary>
        None,
        /// <summary>Font may be embedded, but must be embedded whole — not subsetted.</summary>
        NoSubsetting,
        /// <summary>Font must not be embedded at all (Restricted License).</summary>
        NoEmbedding,
    }
}
