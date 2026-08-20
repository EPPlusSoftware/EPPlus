/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    public sealed class DocumentFontSubsetBuilder
    {
        private readonly OpenTypeFontEngine _engine;
        private readonly SingleFontSubsetter _subsetter = new SingleFontSubsetter();

        // Requested primaries, keyed by request identity. Value carries the primary font instance
        // plus the raw text collected for it (we re-resolve routing in Build, not incrementally).
        private readonly Dictionary<FontKey, RequestedFont> _requested =
            new Dictionary<FontKey, RequestedFont>();

        // ---- Build outputs ----
        private readonly Dictionary<FontKey, OpenTypeFont> _sharedSubsetByIdentity =
            new Dictionary<FontKey, OpenTypeFont>();
        private readonly Dictionary<FontKey, IFontProvider> _providerByRequest =
            new Dictionary<FontKey, IFontProvider>();
        private bool _built;

        public DocumentFontSubsetBuilder(OpenTypeFontEngine engine)
        {
            if (engine == null) throw new ArgumentNullException("engine");
            _engine = engine;
        }

        // ---- Step 1: collect ----
        public void AddText(string family, FontSubFamily subFamily, string text)
        {
            if (_built) throw new InvalidOperationException("Cannot AddText after Build().");
            if (string.IsNullOrEmpty(text)) return;

            var key = new FontKey(family, subFamily);
            RequestedFont req;
            if (!_requested.TryGetValue(key, out req))
            {
                var primary = _engine.LoadFont(family, subFamily);
                req = new RequestedFont(key, primary);
                _requested[key] = req;
            }
            foreach (var cp in CodePointUtil.ExtractCodePoints(text))
                req.CodePoints.Add(cp);
        }

        // ---- Step 2: build ----
        // Add this field alongside the other private fields:
        private readonly Dictionary<FontKey, FontEmbeddingDecision> _decisionByIdentity =
            new Dictionary<FontKey, FontEmbeddingDecision>();

        public void Build()
        {
            if (_built) return;

            var codePointsByIdentity = new Dictionary<FontKey, HashSet<int>>();
            var fontByIdentity = new Dictionary<FontKey, OpenTypeFont>();
            var chainByRequest = new Dictionary<FontKey, List<FontKey>>();

            // ===== PHASE 1: route each code point through the provider, then apply skip =====
            foreach (var kvp in _requested)
            {
                var req = kvp.Value;
                var provider = new DefaultFontProvider(_engine, req.Primary);

                // Distinct destination identities for this request, in first-seen order.
                // First entry becomes the request's primary in phase 3.
                var chainIdentities = new List<FontKey>();

                foreach (var cp in req.CodePoints)
                {
                    // The provider resolves the best font for this code point (primary, or a script-/
                    // emoji-routed fallback), lazy-loading fallbacks as needed.
                    OpenTypeFont dest;
                    ushort glyphId;
                    provider.TryGetGlyphFont((uint)cp, out dest, out glyphId);

                    // If that font may not be embedded, the only replacement in this model is the
                    // last-resort font: the provider yields ONE answer per code point, not a ranked
                    // list, so there is no "next best" to fall to.
                    if (DecisionForFont(dest) == FontEmbeddingDecision.Skip)
                        dest = LastResort();

                    var id = IdentityOf(dest);

                    if (!fontByIdentity.ContainsKey(id))
                        fontByIdentity[id] = dest;

                    HashSet<int> set;
                    if (!codePointsByIdentity.TryGetValue(id, out set))
                        codePointsByIdentity[id] = set = new HashSet<int>();
                    set.Add(cp);

                    if (!chainIdentities.Contains(id))
                        chainIdentities.Add(id);
                }

                // A request with no code points (possible if AddText was called with only skippable
                // content) still needs a primary to shape against.
                if (chainIdentities.Count == 0)
                {
                    var lr = LastResort();
                    var lrId = IdentityOf(lr);
                    if (!fontByIdentity.ContainsKey(lrId))
                        fontByIdentity[lrId] = lr;
                    chainIdentities.Add(lrId);
                }

                chainByRequest[kvp.Key] = chainIdentities;
            }

            // ===== PHASE 2: subset (or embed whole) each identity ONCE =====
            foreach (var kvp in fontByIdentity)
            {
                var id = kvp.Key;
                var font = kvp.Value;

                HashSet<int> cps;
                codePointsByIdentity.TryGetValue(id, out cps);

                if (DecisionForIdentity(id) == FontEmbeddingDecision.EmbedWhole)
                    _sharedSubsetByIdentity[id] = font;
                else
                    _sharedSubsetByIdentity[id] = _subsetter.Subset(font, cps);
            }

            // ===== PHASE 3: build one provider per request from the SHARED subsets =====
            foreach (var kvp in chainByRequest)
            {
                var chain = kvp.Value;
                var provider = new CustomFontProvider(_sharedSubsetByIdentity[chain[0]]);
                for (int i = 1; i < chain.Count; i++)
                    provider.AddFallback(_sharedSubsetByIdentity[chain[i]]);
                _providerByRequest[kvp.Key] = provider;
            }

            _built = true;
        }

        // Loads the last-resort font and ensures a decision is registered for it (it bypasses
        // name resolution, so ResolveEmbeddingDecision is never called for it). It must always be
        // subsettable and must never itself be skipped.
        private OpenTypeFont LastResort()
        {
            var font = EmbeddedFonts.LoadArchivoNarrow(FontSubFamily.Regular);
            _decisionByIdentity[IdentityOf(font)] = FontEmbeddingDecision.Subset;
            return font;
        }

        // Resolves and caches the embedding decision for a font, keyed by identity so the user's
        // OnFontEmbedding hook fires at most once per FontKey. A NoEmbedding font throws here (via
        // ResolveEmbeddingDecision), exactly as in the old per-font path.
        private FontEmbeddingDecision DecisionForFont(OpenTypeFont font)
        {
            var id = IdentityOf(font);
            FontEmbeddingDecision decision;
            if (!_decisionByIdentity.TryGetValue(id, out decision))
            {
                decision = _engine.ResolveEmbeddingDecision(font);
                _decisionByIdentity[id] = decision;
            }
            return decision;
        }

        // Looks up an already-resolved decision by identity. Every identity in fontByIdentity passed
        // through DecisionForFont during phase 1, so it is always present here.
        private FontEmbeddingDecision DecisionForIdentity(FontKey id)
        {
            return _decisionByIdentity[id];
        }

        // Canonical identity from the pre-subset font instance: family + subfamily.
        private static FontKey IdentityOf(OpenTypeFont font)
        {
            return new FontKey(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum());
        }

        /// <summary>
        /// The subsetted fonts to embed — one per distinct font identity used in the document.
        /// Skipped fonts are absent; each shared fallback appears once. Call after Build().
        /// </summary>
        public IEnumerable<SubsettedFont> GetFontsToEmbed()
        {
            RequireBuilt();
            foreach (var kvp in _sharedSubsetByIdentity)
                yield return new SubsettedFont(kvp.Key.Family, kvp.Key.SubFamily, kvp.Value);
        }

        /// <summary>
        /// The provider a given requested font shapes against, wired to the shared subsets.
        /// Returns null if that font was never added. Call after Build().
        /// </summary>
        public IFontProvider GetShapingProvider(string family, FontSubFamily subFamily)
        {
            RequireBuilt();
            IFontProvider provider;
            return _providerByRequest.TryGetValue(new FontKey(family, subFamily), out provider)
                ? provider : null;
        }

        private void RequireBuilt()
        {
            if (!_built)
                throw new InvalidOperationException("Call Build() before reading results.");
        }

        private sealed class RequestedFont
        {
            public FontKey Key { get; private set; }
            public OpenTypeFont Primary { get; private set; }
            public HashSet<int> CodePoints { get; private set; }
            public RequestedFont(FontKey key, OpenTypeFont primary)
            { Key = key; Primary = primary; CodePoints = new HashSet<int>(); }
        }
    }
}