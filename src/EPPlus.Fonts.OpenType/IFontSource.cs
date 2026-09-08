/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2026         EPPlus Software AB           Extracted from OpenTypeFontEngine
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.FontCache
{
    /// <summary>
    /// The font-loading surface a font provider depends on: resolution, availability, and the
    /// configured per-script fallback chains. Nothing else.
    ///
    /// It exists so <see cref="DefaultFontProvider"/> does not depend on
    /// <see cref="OpenTypeFontEngine"/>. A glyph provider has no business reaching the engine's
    /// shaper factory, and while the engine reference was there it could.
    /// </summary>
    internal interface IFontSource
    {
        /// <summary>
        /// The configured fallback chain for a Unicode script, or null if none is configured.
        /// An empty array means fallback is explicitly disabled for that script.
        /// </summary>
        string[] GetScriptFallback(UnicodeScript script);

        /// <summary>
        /// Whether the exact family and subfamily exist, the family exists in another subfamily,
        /// or neither.
        /// </summary>
        FontAvailability GetFontAvailability(string fontName, FontSubFamily subFamily);

        /// <summary>
        /// Resolves, parses and caches a font. Returns null only when the resolver returns null,
        /// which requires a custom <see cref="FontResolver.IFontResolver"/>.
        ///
        /// Deliberately has no ignoreCache parameter, unlike the store's own overload. A script
        /// fallback chain is looked up for every code point in that script, so bypassing the
        /// cache there would be pathological. Leaving the parameter off makes that unavailable
        /// rather than merely discouraged.
        /// </summary>
        OpenTypeFont LoadFont(string fontName, FontSubFamily subFamily);
    }
}