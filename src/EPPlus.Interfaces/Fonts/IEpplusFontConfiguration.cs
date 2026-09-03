/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/27/2026         EPPlus Software AB           Initial implementation
  05/06/2026         EPPlus Software AB           Property-based transactional configuration
  05/20/2026         EPPlus Software AB           Added per-script glyph fallback configuration
 *************************************************************************************************/
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Per-workbook font configuration for EPPlus.
    /// Accessed via <c>ExcelWorkbook.ConfigureFonts(Action&lt;IEpplusFontConfiguration&gt;)</c>.
    /// Changes made inside a ConfigureFonts callback are applied as a single transaction —
    /// when the callback returns, the font resolver is rebuilt and the workbook's font caches are cleared.
    /// </summary>
    public interface IEpplusFontConfiguration
    {
        /// <summary>
        /// Additional directories to search for font files, beyond the system font directories.
        /// Mutate this list inside a Configure callback to add or remove search paths.
        /// </summary>
        IList<string> FontDirectories { get; }

        /// <summary>
        /// Whether the operating system's standard font directories should be searched.
        /// Defaults to <c>true</c>. Set to <c>false</c> to restrict font resolution to the
        /// directories listed in <see cref="FontDirectories"/>.
        /// </summary>
        bool SearchSystemDirectories { get; set; }

        /// <summary>
        /// The font resolver responsible for producing raw TTF/OTF bytes for a requested font.
        /// Set this to replace the default resolver entirely. When a custom resolver is set,
        /// the EPPlus built-in fallback chains and Archivo Narrow ultimate fallback are bypassed —
        /// the resolver is fully responsible for handling missing fonts.
        /// </summary>
        IFontResolver FontResolver { get; set; }

        /// <summary>
        /// User-defined font-name fallback chains.
        /// Each entry maps a font name to an ordered list of fallbacks to try when the primary
        /// is unavailable. Mutate this dictionary inside a Configure callback to add or remove
        /// chains. User chains are tried before the EPPlus built-in fallback chains.
        /// </summary>
        IDictionary<string, string[]> FontFallbacks { get; }

        /// <summary>
        /// Replaces the per-script glyph fallback chain for the given Unicode script.
        /// When the primary font (and any font-level fallbacks) lack a glyph for a character
        /// that belongs to <paramref name="script"/>, EPPlus walks this chain in order and
        /// uses the first font that contains the glyph.
        ///
        /// Setting a chain via this method fully replaces the built-in default for that script.
        /// The caller becomes responsible for providing a complete chain — there is no merge
        /// with the default. Pass an empty array to disable fallback for the script entirely.
        ///
        /// Note: Emoji and Math glyphs are always served by EPPlus's bundled Noto Emoji and
        /// Noto Math fonts respectively, regardless of this configuration.
        /// </summary>
        /// <param name="script">The Unicode script to configure.</param>
        /// <param name="fallbackFontNames">Ordered list of font names to try.</param>
        void SetScriptFallback(UnicodeScript script, params string[] fallbackFontNames);

        /// <summary>
        /// Restores all settings to factory defaults:
        /// <list type="bullet">
        ///   <item>Clears <see cref="FontDirectories"/>.</item>
        ///   <item>Sets <see cref="SearchSystemDirectories"/> to <c>true</c>.</item>
        ///   <item>Restores the default <see cref="FontResolver"/> (with Archivo Narrow built-in fallback).</item>
        ///   <item>Clears <see cref="FontFallbacks"/>.</item>
        ///   <item>Restores the default per-script glyph fallback chains.</item>
        /// </list>
        /// </summary>
    
        /// <summary>
        /// Whether text measurement may fall back to serialized font metrics when the requested
        /// font is not available as a font file. Defaults to
        /// <see cref="MetricsFallbackMode.WhenFontMissing"/>.
        ///
        /// Measurement and line breaking only. Rendering and font embedding always require a real
        /// font file, so PDF export is unaffected by this setting.
        /// </summary>
        MetricsFallbackMode MetricsFallback { get; set; }

        void Reset();

        /// <summary>
        /// Registers a callback invoked for each font that is about to be embedded, letting the
        /// caller override how EPPlus handles the font's declared embedding restriction (fsType).
        /// Return <see cref="FontEmbeddingDecision.Default"/> to keep EPPlus's standard behaviour.
        /// </summary>
        /// <remarks>
        /// A font may declare that it must not be embedded (Restricted License) or must not be
        /// subsetted. By returning <see cref="FontEmbeddingDecision.Subset"/> or
        /// <see cref="FontEmbeddingDecision.EmbedWhole"/>, the caller asserts they hold the rights
        /// to do so; EPPlus cannot verify any licence the caller may have obtained from the font's
        /// owner. Only one callback is active; a later call replaces the earlier one.
        /// </remarks>
        void OnFontEmbedding(Func<FontEmbeddingInfo, FontEmbeddingDecision> callback);
    }

}