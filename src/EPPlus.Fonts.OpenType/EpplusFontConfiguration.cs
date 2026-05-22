/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB.
  This software is licensed under PolyForm Noncommercial License 1.0.0
  and may only be used for noncommercial purposes
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  02/27/2026         EPPlus Software AB           Replaces FontResolutionConfig
  05/06/2026         EPPlus Software AB           Property-based transactional configuration
  05/20/2026         EPPlus Software AB           Added per-script glyph fallback configuration
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.FontResolver
{
    /// <summary>
    /// Concrete implementation of <see cref="IEpplusFontConfiguration"/>.
    /// Managed exclusively by <see cref="OpenTypeFonts"/> — not instantiated by user code.
    /// Mutations are intended to happen inside an <c>OpenTypeFonts.Configure</c> callback,
    /// after which <see cref="OpenTypeFonts"/> reads the resulting state as a snapshot and
    /// rebuilds the resolver.
    /// </summary>
    internal class EpplusFontConfiguration : IEpplusFontConfiguration
    {
        private readonly List<string> _fontDirectories = new List<string>();
        private readonly Dictionary<string, string[]> _fontFallbacks =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<UnicodeScript, string[]> _scriptFallbacks =
            new Dictionary<UnicodeScript, string[]>();

        public EpplusFontConfiguration()
        {
            SearchSystemDirectories = true;
            ApplyDefaultScriptFallbacks();
        }

        /// <inheritdoc/>
        public IList<string> FontDirectories
        {
            get { return _fontDirectories; }
        }

        /// <inheritdoc/>
        public bool SearchSystemDirectories { get; set; }

        /// <inheritdoc/>
        public IFontResolver FontResolver { get; set; }

        /// <inheritdoc/>
        public IDictionary<string, string[]> FontFallbacks
        {
            get { return _fontFallbacks; }
        }

        /// <inheritdoc/>
        public void SetScriptFallback(UnicodeScript script, params string[] fallbackFontNames)
        {
            if (fallbackFontNames == null)
            {
                _scriptFallbacks[script] = new string[0];
                return;
            }

            // Copy to insulate against caller mutating the array after the call.
            var copy = new string[fallbackFontNames.Length];
            Array.Copy(fallbackFontNames, copy, fallbackFontNames.Length);
            _scriptFallbacks[script] = copy;
        }

        /// <inheritdoc/>
        public void Reset()
        {
            _fontDirectories.Clear();
            SearchSystemDirectories = true;
            FontResolver = null;
            _fontFallbacks.Clear();
            _scriptFallbacks.Clear();
            ApplyDefaultScriptFallbacks();
        }

        // -----------------------------------------------------------------------------------------
        // Internal API — consumed by DefaultFontResolver and DefaultFontProvider
        // in the same assembly.
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Returns the user-configured fallback chain for the given font name, or null if none
        /// is configured. Case-insensitive lookup.
        /// </summary>
        internal string[] GetFallbacks(string fontName)
        {
            string[] result;
            return _fontFallbacks.TryGetValue(fontName, out result) ? result : null;
        }

        /// <summary>
        /// Returns the configured fallback chain for the given Unicode script, or null if
        /// none is configured. A returned empty array means fallback is explicitly disabled
        /// for the script.
        /// </summary>
        internal string[] GetScriptFallback(UnicodeScript script)
        {
            string[] result;
            return _scriptFallbacks.TryGetValue(script, out result) ? result : null;
        }

        // -----------------------------------------------------------------------------------------
        // Default per-script chains
        //
        // These cover the scripts most likely to appear in Office documents. Each chain prefers
        // platform-native fonts first (Windows / macOS), then Noto as a Linux / open-source
        // fallback. Chains stay within a single language family — falling back from Japanese
        // to Chinese, or vice versa, would render incorrect glyph forms for shared CJK
        // ideographs.
        //
        // Emoji and Math are intentionally NOT in this table. They are served by EPPlus's
        // bundled Noto Emoji and Noto Math fonts and are routed before per-script lookup.
        // -----------------------------------------------------------------------------------------

        private void ApplyDefaultScriptFallbacks()
        {
            _scriptFallbacks[UnicodeScript.Han] = new[]
            {
                "Microsoft YaHei", "SimSun", "Noto Sans CJK SC", "PingFang SC"
            };

            _scriptFallbacks[UnicodeScript.Hiragana] = new[]
            {
                "Yu Gothic", "MS Gothic", "Meiryo", "Noto Sans CJK JP"
            };

            _scriptFallbacks[UnicodeScript.Katakana] = new[]
            {
                "Yu Gothic", "MS Gothic", "Meiryo", "Noto Sans CJK JP"
            };

            _scriptFallbacks[UnicodeScript.Hangul] = new[]
            {
                "Malgun Gothic", "Gulim", "Noto Sans CJK KR"
            };

            _scriptFallbacks[UnicodeScript.Arabic] = new[]
            {
                "Segoe UI", "Tahoma", "Arial", "Noto Sans Arabic"
            };

            _scriptFallbacks[UnicodeScript.Hebrew] = new[]
            {
                "Segoe UI", "Tahoma", "Arial", "Noto Sans Hebrew"
            };

            _scriptFallbacks[UnicodeScript.Thai] = new[]
            {
                "Tahoma", "Leelawadee UI", "Noto Sans Thai"
            };

            _scriptFallbacks[UnicodeScript.Devanagari] = new[]
            {
                "Mangal", "Nirmala UI", "Noto Sans Devanagari"
            };
        }
    }
}