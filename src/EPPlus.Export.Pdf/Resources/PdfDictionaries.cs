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
 *************************************************************************************************/
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Integration;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Export.Pdf.Settings;

namespace EPPlus.Export.Pdf.Resources
{
    internal class PdfDictionaries
    {
        internal readonly Dictionary<FontKey, PdfFontResource> Fonts = new Dictionary<FontKey, PdfFontResource>();
        internal readonly Dictionary<string, PdfPatternResource> Patterns = new Dictionary<string, PdfPatternResource>();
        internal readonly Dictionary<string, PdfShadingResource> Shadings = new Dictionary<string, PdfShadingResource>();
        internal Dictionary<FontKey, IFontProvider> ShapedProviders = new Dictionary<FontKey, IFontProvider>();

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

        public void AddFont(PdfPageSettings pageSettings, string FontName, FontSubFamily SubFamily, string Text)
        {
            var key = ResolveFontKey(pageSettings, FontName, SubFamily);
            if (!Fonts.ContainsKey(key))
            {
                int label = 1;
                if (Fonts.Count > 0)
                {
                    label = Fonts.Last().Value.labelNumber + 1;
                }
                Fonts.Add(key, new PdfFontResource(FontName, SubFamily, label, pageSettings));
            }
            var manger = Fonts[key].fontSubsetManager;
            manger.AddText(Text);
        }

        internal PdfFontResource GetFont(PdfPageSettings pageSettings, string fontName, FontSubFamily subFamily)
        {
            var key = ResolveFontKey(pageSettings, fontName, subFamily);
            if (!Fonts.ContainsKey(key))
            {
                throw new KeyNotFoundException("Font: " + key + " is missing from dictionary.");
            }
            return Fonts[key];
        }
    }
}