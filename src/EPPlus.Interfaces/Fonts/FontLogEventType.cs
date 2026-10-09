/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/08/2026         EPPlus Software AB           Font logging
 *************************************************************************************************/
namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// The kind of decision a <see cref="FontLogEvent"/> describes.
    /// </summary>
    public enum FontLogEventType
    {
        /// <summary>A requested font was found as an exact match.</summary>
        FontResolved,
        /// <summary>A requested font was replaced by an entry in a user-configured or built-in fallback chain.</summary>
        FontFallback,
        /// <summary>No exact match or chain entry was found; the embedded last-resort font is used.</summary>
        FontLastResort,
        /// <summary>The font resolver returned no font for the request.</summary>
        FontNotResolved,
        /// <summary>A font was loaded from the resolver (cache miss). Reports a substitution made by a custom resolver.</summary>
        FontLoaded,
        /// <summary>Text is measured from serialized font metrics instead of a font file.</summary>
        MeasurementMetricsUsed,
        /// <summary>A font is replaced by a web substitute for <see cref="FontRenderTarget.Web"/>.</summary>
        WebFontSubstitution,
        /// <summary>The embedding policy was decided for a font.</summary>
        EmbeddingDecision,
        /// <summary>The fallback chain of a script was resolved, with the outcome per entry.</summary>
        ScriptChainResolved,
        /// <summary>A font in a script fallback chain could not be loaded.</summary>
        ScriptFontLoadFailed,
        /// <summary>A fallback font supplied a glyph for the first time.</summary>
        ScriptFallbackUsed,
        /// <summary>A code point was routed to a fallback font (reported once per code point).</summary>
        GlyphFallback,
        /// <summary>No candidate font has a glyph for a code point (reported once per code point).</summary>
        GlyphMissing
    }
}