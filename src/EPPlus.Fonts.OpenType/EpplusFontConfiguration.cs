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

        public EpplusFontConfiguration()
        {
            SearchSystemDirectories = true;
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
        public void Reset()
        {
            _fontDirectories.Clear();
            SearchSystemDirectories = true;
            FontResolver = null;
            _fontFallbacks.Clear();
        }

        // -----------------------------------------------------------------------------------------
        // Internal API — consumed by DefaultFontResolver in the same assembly.
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
    }
}