/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
  08/17/2026         EPPlus Software AB           Canonical FontKey + resolve cache
  08/20/2026         EPPlus Software AB           Document-wide subsetting via DocumentFontSubsetBuilder
 *************************************************************************************************/
using EPPlus.Export.Pdf.Settings;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Subsetting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Export.Pdf.Resources
{
    internal class PdfDictionaries
    {
        internal readonly Dictionary<FontKey, PdfFontResource> Fonts = new Dictionary<FontKey, PdfFontResource>();
        internal readonly Dictionary<string, PdfPatternResource> Patterns = new Dictionary<string, PdfPatternResource>();
        internal readonly Dictionary<string, PdfShadingResource> Shadings = new Dictionary<string, PdfShadingResource>();
        internal Dictionary<FontKey, IFontProvider> ShapedProviders = new Dictionary<FontKey, IFontProvider>();

        // One document-wide subset builder, replacing the per-font FontSubsetManager. Owns all
        // fallback resolution, embedding-restriction decisions, and shared subset construction.
        private DocumentFontSubsetBuilder _subsetBuilder;

        // Cache mapping a requested (family, subfamily) to the canonical FontKey of
        // the loaded font. Case-insensitive on the requested family so casing in the
        // source workbook resolves to the same key. Ensures the font is only loaded
        // once per distinct request.
        private readonly Dictionary<string, FontKey> _requestedToKey =
            new Dictionary<string, FontKey>();

        private static string BuildRequestCacheKey(string family, FontSubFamily subFamily)
        {
            // Lower-case the requested family for case-insensitive lookup; subfamily
            // is an enum so its numeric value is stable.
            string fam = family == null ? string.Empty : family.ToLowerInvariant();
            return fam + "|" + ((int)subFamily);
        }

        /// <summary>
        /// Resolves a requested (family, subfamily) to the canonical FontKey of the
        /// loaded font, loading the font at most once per distinct request.
        /// </summary>
        internal FontKey ResolveFontKey(PdfPageSettings pageSettings, string family, FontSubFamily subFamily)
        {
            var cacheKey = BuildRequestCacheKey(family, subFamily);
            FontKey key;
            if (_requestedToKey.TryGetValue(cacheKey, out key))
            {
                return key;
            }

            var font = pageSettings.FontEngine.LoadFont(family, subFamily);
            key = new FontKey(font.GetEnglishFontFamilyName(), font.NameTable.GetSubfamilyEnum());
            _requestedToKey[cacheKey] = key;
            return key;
        }

        // CHANGE 1: AddFont now only feeds the builder. It no longer creates a PdfFontResource —
        // resources are created later, per ACTUAL font, during shaping. We still resolve the
        // requested key so it is registered in _requestedToKey for later provider wiring.
        public void AddFont(PdfPageSettings pageSettings, string fontName, FontSubFamily subFamily, string text)
        {
            EnsureBuilder(pageSettings);
            ResolveFontKey(pageSettings, fontName, subFamily);   // register the requested key
            _subsetBuilder.AddText(fontName, subFamily, text);
        }

        private void EnsureBuilder(PdfPageSettings pageSettings)
        {
            if (_subsetBuilder == null)
                _subsetBuilder = new DocumentFontSubsetBuilder(pageSettings.FontEngine);
        }

        // CHANGE 2: new. Runs the single document-wide build, then wires one shaping provider per
        // requested font. Call once, after all text is collected, before shaping. Replaces the
        // old per-font CreateSubsettedProvider loop in PdfCatalog.
        internal void BuildSubsets(PdfPageSettings pageSettings)
        {
            if (_subsetBuilder == null) return;   // no text was collected
            _subsetBuilder.Build();

            foreach (var requestedKey in _requestedToKey.Values.Distinct())
            {
                var provider = _subsetBuilder.GetShapingProvider(requestedKey.Family, requestedKey.SubFamily);
                if (provider != null)
                    ShapedProviders[requestedKey] = provider;
            }
        }

        // CHANGE 3: GetFont is used by the renderer for METRICS only (glyph font selection is done
        // per-glyph via FontIdMap). After skipping, the requested font may not be embedded, so we
        // translate the requested font to the ACTUAL primary that renders it (the shaping
        // provider's primary) and return that resource.
        internal PdfFontResource GetFont(PdfPageSettings pageSettings, string fontName, FontSubFamily subFamily)
        {
            var requestedKey = ResolveFontKey(pageSettings, fontName, subFamily);

            // Preferred path: translate requested -> actual via the shaping provider's primary.
            IFontProvider provider;
            if (ShapedProviders.TryGetValue(requestedKey, out provider) && provider.PrimaryFont != null)
            {
                var actual = provider.PrimaryFont;
                var actualKey = new FontKey(actual.GetEnglishFontFamilyName(), actual.NameTable.GetSubfamilyEnum());
                PdfFontResource viaProvider;
                if (Fonts.TryGetValue(actualKey, out viaProvider))
                    return viaProvider;
            }

            // Fallback: the requested font was embedded under its own identity (not skipped).
            PdfFontResource direct;
            if (Fonts.TryGetValue(requestedKey, out direct))
                return direct;

            throw new KeyNotFoundException("Font: " + requestedKey + " is missing from dictionary.");
        }
    }
}