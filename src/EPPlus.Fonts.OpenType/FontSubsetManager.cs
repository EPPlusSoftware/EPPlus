/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/25/2026         EPPlus Software AB           Font subset manager for PDF export
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Prepares subsetted fonts for PDF export by pre-scanning text,
    /// distributing code points to the correct font via the fallback chain,
    /// and creating minimal subsets of all fonts (including fallbacks).
    /// 
    /// Usage:
    ///   1. Create with an IFontProvider (e.g., DefaultFontProvider)
    ///   2. Call AddText() for all text that will be rendered (e.g., all cell values)
    ///   3. Call CreateSubsettedProvider() to get a new IFontProvider with subsetted fonts
    ///   4. Use the returned provider for shaping and PDF rendering
    /// </summary>
    public class FontSubsetManager
    {
        private readonly IFontProvider _sourceProvider;

        // Code points collected per font (key = original font instance)
        private readonly Dictionary<OpenTypeFont, HashSet<int>> _codePointsByFont =
            new Dictionary<OpenTypeFont, HashSet<int>>();

        public FontSubsetManager(IFontProvider sourceProvider)
        {
            if (sourceProvider == null)
                throw new ArgumentNullException("sourceProvider");

            _sourceProvider = sourceProvider;
        }

        public FontSubsetManager(OpenTypeFontEngine engine, OpenTypeFont font)
            : this(new DefaultFontProvider(engine, font))
        {
            
        }

        /// <summary>
        /// Scans text and distributes each code point to the font that will render it.
        /// Call this for every piece of text that will appear in the document.
        /// </summary>
        public void AddText(string text)
        {
            if (string.IsNullOrEmpty(text))
                return;

            var codePoints = CodePointUtil.ExtractCodePoints(text);

            foreach (var cp in codePoints)
            {
                OpenTypeFont font;
                ushort glyphId;
                _sourceProvider.TryGetGlyphFont((uint)cp, out font, out glyphId);

                HashSet<int> fontCodePoints;
                if (!_codePointsByFont.TryGetValue(font, out fontCodePoints))
                {
                    fontCodePoints = new HashSet<int>();
                    _codePointsByFont[font] = fontCodePoints;
                }

                fontCodePoints.Add(cp);
            }
        }

        /// <summary>
        /// Creates a new IFontProvider where all fonts (primary + fallbacks) are subsetted
        /// to contain only the glyphs needed for the collected text.
        /// Fonts that had no text collected are excluded from the result.
        /// </summary>
        public IFontProvider CreateSubsettedProvider()
        {
            var primaryFont = _sourceProvider.PrimaryFont;
            var allFonts = _sourceProvider.GetAllFonts().ToList();

            // Subset each font that has collected code points
            var subsetMap = new Dictionary<OpenTypeFont, OpenTypeFont>();

            foreach (var kvp in _codePointsByFont)
            {
                var originalFont = kvp.Key;
                var codePoints = kvp.Value;

                if (codePoints.Count == 0)
                    continue;

                try
                {
                    var chars = CodePointUtil.CodePointsToString(codePoints);
                    var subset = originalFont.CreateSubset(chars);
                    subsetMap[originalFont] = subset;
                }
                catch (Exception ex)
                {
                    // If subsetting fails, use the original font
                    System.Diagnostics.Debug.WriteLine(
                        $"Warning: Could not subset '{originalFont.NameTable?.GetFullFontName()}': {ex.Message}");
                    subsetMap[originalFont] = originalFont;
                }
            }

            // Build new provider with subsetted fonts, preserving fallback order
            var subsetPrimary = subsetMap.ContainsKey(primaryFont)
                ? subsetMap[primaryFont]
                : primaryFont;

            var provider = new CustomFontProvider(subsetPrimary);

            // Add fallback fonts in their original order (skip primary)
            for (int i = 1; i < allFonts.Count; i++)
            {
                var originalFallback = allFonts[i];

                if (subsetMap.ContainsKey(originalFallback))
                {
                    provider.AddFallback(subsetMap[originalFallback]);
                }
                // If no code points were collected for this fallback, skip it entirely
            }

            return provider;
        }
    }
}