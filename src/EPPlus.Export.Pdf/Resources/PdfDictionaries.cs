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
        internal readonly Dictionary<string, PdfImageResource> Images = new Dictionary<string, PdfImageResource>();
        internal Dictionary<FontKey, IFontProvider> ShapedProviders = new Dictionary<FontKey, IFontProvider>();
        private DocumentFontSubsetBuilder _subsetBuilder;
        private readonly Dictionary<string, FontKey> _requestedToKey = new Dictionary<string, FontKey>();

        private static string BuildRequestCacheKey(string family, FontSubFamily subFamily)
        {
            string fam = family == null ? string.Empty : family.ToLowerInvariant();
            return fam + "|" + ((int)subFamily);
        }

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

        public void AddFont(PdfPageSettings pageSettings, string fontName, FontSubFamily subFamily, string text)
        {
            EnsureBuilder(pageSettings);
            ResolveFontKey(pageSettings, fontName, subFamily);
            _subsetBuilder.AddText(fontName, subFamily, text);
        }

        private void EnsureBuilder(PdfPageSettings pageSettings)
        {
            if (_subsetBuilder == null)
                _subsetBuilder = new DocumentFontSubsetBuilder(pageSettings.FontEngine);
        }

        internal void BuildSubsets(PdfPageSettings pageSettings)
        {
            if (_subsetBuilder == null) return;
            _subsetBuilder.Build();

            foreach (var requestedKey in _requestedToKey.Values.Distinct())
            {
                var provider = _subsetBuilder.GetShapingProvider(requestedKey.Family, requestedKey.SubFamily);
                if (provider != null)
                    ShapedProviders[requestedKey] = provider;
            }
        }

        internal PdfFontResource GetFont(PdfPageSettings pageSettings, string fontName, FontSubFamily subFamily)
        {
            var requestedKey = ResolveFontKey(pageSettings, fontName, subFamily);
            IFontProvider provider;
            if (ShapedProviders.TryGetValue(requestedKey, out provider) && provider.PrimaryFont != null)
            {
                var actual = provider.PrimaryFont;
                var actualKey = new FontKey(actual.GetEnglishFontFamilyName(), actual.NameTable.GetSubfamilyEnum());
                PdfFontResource viaProvider;
                if (Fonts.TryGetValue(actualKey, out viaProvider))
                    return viaProvider;
            }
            // Fallback
            PdfFontResource direct;
            if (Fonts.TryGetValue(requestedKey, out direct))
                return direct;
            throw new KeyNotFoundException("Font: " + requestedKey + " is missing from dictionary.");
        }

        internal PdfImageResource AddImage(byte[] imageBytes)
        {
            var key = GetImageKey(imageBytes);
            if (!Images.TryGetValue(key, out var res))
            {
                int label = 1;
                if (Images.Count > 0) label = Images.Last().Value.labelNumber + 1;
                res = new PdfImageResource(label, imageBytes);
                Images.Add(key, res);
            }
            return res;
        }

        private static string GetImageKey(byte[] bytes)
        {
            using (var sha = System.Security.Cryptography.SHA1.Create())
            {
                return System.Convert.ToBase64String(sha.ComputeHash(bytes));
            }
        }
    }
}