/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/02/2026         EPPlus Software AB           Extracted from OpenTypeFontEngine
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.FontCache
{
    /// <summary>
    /// Resolves, parses and caches fonts for one <see cref="OpenTypeFontEngine"/>.
    /// Owns the resolver, the parsed-font cache and the per-font locks.
    ///
    /// One instance per engine: two engines never share parsed fonts, because their resolver
    /// configurations may produce different fonts for the same name.
    /// </summary>
    internal class FontStore : IFontSource
    {
        private readonly object _syncRoot = new object();
        private readonly Dictionary<string, object> _fontLocks = new Dictionary<string, object>();
        private readonly OpenTypeFontCache _fontCache = new OpenTypeFontCache();
        private readonly IFontResolver _resolver;
        private readonly EpplusFontConfiguration _configuration;

        private bool _disposed;

        internal FontStore(IFontResolver resolver, EpplusFontConfiguration configuration)
        {
            if (resolver == null)
                throw new ArgumentNullException("resolver");
            if (configuration == null)
                throw new ArgumentNullException("configuration");

            _resolver = resolver;
            _configuration = configuration;
        }

        // -----------------------------------------------------------------------------------------
        // Font loading
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Loads a font by name and subfamily, with thread-safe caching.
        /// Returns null if the font cannot be resolved.
        /// </summary>
        internal OpenTypeFont LoadFont(string fontName, FontSubFamily subFamily, bool ignoreCache)
        {
            ThrowIfDisposed();

            if (ignoreCache)
                return ResolveAndCreate(_resolver, fontName, subFamily);

            string lockKey = BuildCacheKey(fontName, subFamily);
            object fontLock;
            lock (_syncRoot)
            {
                if (!_fontLocks.TryGetValue(lockKey, out fontLock))
                {
                    fontLock = new object();
                    _fontLocks[lockKey] = fontLock;
                }
            }

            lock (fontLock)
            {
                var cached = _fontCache.GetFromCache(lockKey);
                if (cached != null && cached.Font != null && cached.IsLoaded)
                {
                    cached.Font.EnsureFullyLoaded();
                    return cached.Font;
                }

                _fontCache.BeginCache(lockKey);

                var font = ResolveAndCreate(_resolver, fontName, subFamily);
                if (font == null)
                {
                    // BeginCache left a not-loaded placeholder. Nothing will ever complete it,
                    // so remove it — otherwise every later GetFromCache for this key spends the
                    // full two-second Monitor.Wait timeout before giving up.
                    _fontCache.RemoveIfNotLoaded(lockKey);
                    return null;
                }

                font.EnsureFullyLoaded();
                font.IsReadOnly = true;
                _fontCache.AddToCache(font, lockKey);
                return font;
            }
        }

        /// <inheritdoc/>
        public OpenTypeFont LoadFont(string fontName, FontSubFamily subFamily)
        {
            return LoadFont(fontName, subFamily, false);
        }

        // -----------------------------------------------------------------------------------------
        // Availability and configuration
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Checks whether a font is available in the configured font system.
        ///
        /// If the resolver implements <see cref="IFontAvailabilityProvider"/> the call delegates
        /// to it. Otherwise it probes via <see cref="IFontResolver.ResolveFont"/>, which can only
        /// distinguish found from not found — never <see cref="FontAvailability.FamilyOnly"/>, and
        /// never NotFound at all for a resolver that substitutes internally.
        /// </summary>
        public FontAvailability GetFontAvailability(string fontName, FontSubFamily subFamily)
        {
            ThrowIfDisposed();
            if (fontName == null)
                throw new ArgumentNullException("fontName");

            var provider = _resolver as IFontAvailabilityProvider;
            if (provider != null)
                return provider.GetFontAvailability(fontName, subFamily);

            return _resolver.ResolveFont(fontName, subFamily) != null
                ? FontAvailability.Exact
                : FontAvailability.NotFound;
        }

        /// <inheritdoc/>
        public string[] GetScriptFallback(UnicodeScript script)
        {
            return _configuration.GetScriptFallback(script);
        }

        // -----------------------------------------------------------------------------------------
        // Lifecycle
        // -----------------------------------------------------------------------------------------

        internal void Clear()
        {
            lock (_syncRoot)
            {
                _fontCache.Clear();
                _fontLocks.Clear();
            }
        }

        /// <summary>
        /// Called by the engine from Dispose. A <see cref="DefaultFontProvider"/> holds this store
        /// directly, so without a flag here it could keep loading fonts after the engine that owns
        /// it was disposed — the engine's own disposed check no longer covers every path in.
        /// </summary>
        internal void MarkDisposed()
        {
            _disposed = true;
            Clear();
        }

        private void ThrowIfDisposed()
        {
            if (_disposed)
                throw new ObjectDisposedException("OpenTypeFontEngine");
        }

        // -----------------------------------------------------------------------------------------
        // Helpers
        // -----------------------------------------------------------------------------------------

        internal static string BuildCacheKey(string fontName, FontSubFamily subFamily)
        {
            return string.Format("{0}_{1}", fontName, subFamily);
        }

        private static OpenTypeFont ResolveAndCreate(IFontResolver resolver, string fontName, FontSubFamily subFamily)
        {
            var bytes = resolver.ResolveFont(fontName, subFamily);
            if (bytes == null)
                return null;

            return new OpenTypeFont(bytes);
        }
    }
}