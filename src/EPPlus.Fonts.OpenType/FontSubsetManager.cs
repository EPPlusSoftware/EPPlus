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
using OfficeOpenXml.Interfaces.Fonts;
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
        private readonly OpenTypeFontEngine _fontEngine;

        // Code points collected per font (key = original font instance)
        private readonly Dictionary<OpenTypeFont, HashSet<int>> _codePointsByFont =
            new Dictionary<OpenTypeFont, HashSet<int>>();

        public FontSubsetManager(OpenTypeFontEngine engine, IFontProvider sourceProvider)
        {
            if (engine == null)
                throw new ArgumentNullException("engine");
            if (sourceProvider == null)
                throw new ArgumentNullException("sourceProvider");

            _sourceProvider = sourceProvider;
            _fontEngine = engine;
        }

        public FontSubsetManager(OpenTypeFontEngine engine, OpenTypeFont font)
            : this(engine, new DefaultFontProvider(engine, font))
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

                //var fontName = font?.NameTable?.GetFullFontName() ?? "null";
                //if ((char)cp == 'E' || (char)cp == 'P')
                //{
                //    Console.WriteLine($"[FontSubsetManager.AddText] cp='{(char)cp}' (U+{cp:X4}) -> font='{fontName}', glyphId={glyphId}");
                //}

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
            var originalChain = _sourceProvider.GetAllFonts().ToList();

            // --- Step 1: chain-level decision. Call ResolveEmbeddingDecision ONCE per font,
            // outside try/catch (a NoEmbedding font must throw straight to the caller). ---
            var decisions = new Dictionary<OpenTypeFont, FontEmbeddingDecision>();
            var effectiveChain = new List<OpenTypeFont>();   // ordered, skipped fonts removed
            foreach (var font in originalChain)
            {
                var decision = _fontEngine.ResolveEmbeddingDecision(font);
                decisions[font] = decision;
                if (decision != FontEmbeddingDecision.Skip)
                    effectiveChain.Add(font);
            }

            // If everything was skipped, pull in the last-resort font so the chain is never empty.
            if (effectiveChain.Count == 0)
                effectiveChain.Add(EmbeddedFonts.LoadArchivoNarrow(FontSubFamily.Regular));

            // --- Step 2: redistribute the skipped fonts' code points over the reduced chain. ---
            foreach (var font in originalChain)
            {
                if (decisions[font] != FontEmbeddingDecision.Skip)
                    continue;

                HashSet<int> cps;
                if (_codePointsByFont.TryGetValue(font, out cps))
                {
                    foreach (var cp in cps)
                    {
                        var target = ResolveOverChain(effectiveChain, cp); // cmap walk, ultimately chain[0]
                        HashSet<int> targetCps;
                        if (!_codePointsByFont.TryGetValue(target, out targetCps))
                            _codePointsByFont[target] = targetCps = new HashSet<int>();
                        targetCps.Add(cp);
                    }
                }
                _codePointsByFont.Remove(font);  // a skipped font is never subsetted
            }

            // --- Step 3: subset loop, now only over fonts in effectiveChain.
            // Same switch as before BUT the Skip branch is gone — it can no longer occur here. ---
            var subsetMap = new Dictionary<OpenTypeFont, OpenTypeFont>();
            foreach (var font in effectiveChain)
            {
                HashSet<int> cps;
                if (!_codePointsByFont.TryGetValue(font, out cps) || cps.Count == 0)
                    continue;

                switch (decisions.ContainsKey(font) ? decisions[font] : FontEmbeddingDecision.Subset)
                {
                    case FontEmbeddingDecision.EmbedWhole:
                        subsetMap[font] = font;
                        break;
                    case FontEmbeddingDecision.Subset:
                        try { subsetMap[font] = font.CreateSubset(CodePointUtil.CodePointsToString(cps)); }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine(
                                $"Warning: could not subset '{font.NameTable?.GetFullFontName()}': {ex.Message}");
                            subsetMap[font] = font;
                        }
                        break;
                }
            }

            // --- Step 4: build the provider. effectiveChain[0] becomes the primary — a skipped
            // primary is already filtered out, so "primary is replaced" is expressed naturally. ---
            var provider = new CustomFontProvider(Resolved(effectiveChain[0], subsetMap));
            for (int i = 1; i < effectiveChain.Count; i++)
                provider.AddFallback(Resolved(effectiveChain[i], subsetMap));
            return provider;
        }

        private static OpenTypeFont Resolved(OpenTypeFont font, Dictionary<OpenTypeFont, OpenTypeFont> map)
        {
            // A font with no collected code points is kept unchanged.
            OpenTypeFont subset;
            return map.TryGetValue(font, out subset) ? subset : font;
        }

        // Chain-local cmap lookup. Last resort: chain[0] (which, in the all-skipped case, IS Archivo Narrow).
        private static OpenTypeFont ResolveOverChain(List<OpenTypeFont> chain, int codePoint)
        {
            foreach (var font in chain)
            {
                ushort glyphId;
                if (font.CmapTable.TryGetGlyphId((uint)codePoint, out glyphId))
                    return font;
            }
            return chain[0];
        }
    }
}