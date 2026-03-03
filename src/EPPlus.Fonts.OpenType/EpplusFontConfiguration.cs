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
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.FontResolver
{
    /// <summary>
    /// Concrete implementation of <see cref="IEpplusFontConfiguration"/>.
    /// Managed exclusively by <see cref="OpenTypeFonts"/> — not instantiated by user code.
    /// </summary>
    internal class EpplusFontConfiguration : IEpplusFontConfiguration
    {
        private readonly Dictionary<string, string[]> _fallbacks =
            new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);

        // Internal events consumed by OpenTypeFonts in the same assembly.
        internal event Action OnReset;
        internal event Action<IFontResolver> OnSetFontResolver;

        // -----------------------------------------------------------------------------------------
        // IEpplusFontConfiguration
        // -----------------------------------------------------------------------------------------

        /// <inheritdoc/>
        public IEpplusFontConfiguration AddFallback(string primaryFont, params string[] fallbacks)
        {
            if (string.IsNullOrEmpty(primaryFont))
                throw new ArgumentNullException("primaryFont");
            if (fallbacks == null || fallbacks.Length == 0)
                throw new ArgumentException("At least one fallback must be specified.", "fallbacks");

            // Additive: merge with any existing fallbacks for this font.
            string[] existing;
            if (_fallbacks.TryGetValue(primaryFont, out existing))
            {
                var merged = new string[existing.Length + fallbacks.Length];
                existing.CopyTo(merged, 0);
                fallbacks.CopyTo(merged, existing.Length);
                _fallbacks[primaryFont] = merged;
            }
            else
            {
                _fallbacks[primaryFont] = fallbacks;
            }

            return this;
        }

        /// <inheritdoc/>
        public IEpplusFontConfiguration SetFontResolver(IFontResolver resolver)
        {
            if (resolver == null)
                throw new ArgumentNullException("resolver");

            if (OnSetFontResolver != null)
                OnSetFontResolver(resolver);

            return this;
        }

        /// <inheritdoc/>
        public IEpplusFontConfiguration Reset()
        {
            _fallbacks.Clear();

            if (OnReset != null)
                OnReset();

            return this;
        }

        // -----------------------------------------------------------------------------------------
        // Internal API — consumed by DefaultFontResolver in the same assembly.
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Returns the fallback chain for the given font name, or null if none is configured.
        /// </summary>
        internal string[] GetFallbacks(string fontName)
        {
            string[] result;
            return _fallbacks.TryGetValue(fontName, out result) ? result : null;
        }
    }
}