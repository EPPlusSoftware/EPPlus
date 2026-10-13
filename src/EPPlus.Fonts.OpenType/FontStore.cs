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
  10/08/2026         EPPlus Software AB           Font logging
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Logging;
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

        // DefaultFontResolver logs the reason for each decision itself. A custom resolver does
        // not, so for those the store reports a substitution it can observe from the outside.
        private readonly bool _resolverExplainsItself;

        private bool _disposed;

        internal FontStore(IFontResolver resolver, EpplusFontConfiguration configuration)
        {
            if (resolver == null)
                throw new ArgumentNullException("resolver");
            if (configuration == null)
                throw new ArgumentNullException("configuration");

            _resolver = resolver;
            _configuration = configuration;
            _resolverExplainsItself = resolver is DefaultFontResolver;
        }

        /// <inheritdoc/>
        public IFontLogger Logger
        {
            get { return _configuration.ActiveLogger; }
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
            {
                var uncached = ResolveAndCreate(_resolver, fontName, subFamily);
                LogLoaded(fontName, subFamily, uncached);
                return uncached;
            }

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

                // Only reached on a cache miss, so this reports each font once per engine.
                // Held under the per-font lock, which only blocks other loads of the same font.
                LogLoaded(fontName, subFamily, font);

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
        // Logging
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// Reports the outcome of a resolver call. An unresolved font is a warning. A resolved font
        /// is debug output, except when a custom resolver returned a different family than was
        /// asked for, which nothing else would explain.
        /// </summary>
        private void LogLoaded(string fontName, FontSubFamily subFamily, OpenTypeFont font)
        {
            var logger = Logger;

            if (font == null)
            {
                if (FontLog.IsEnabled(logger, FontLogSeverity.Warning))
                {
                    FontLog.Write(
                        logger,
                        FontLogSeverity.Warning,
                        FontLogEventType.FontNotResolved,
                        string.Format("Requested font '{0}' {1} could not be resolved: the font resolver returned no font.", fontName, subFamily),
                        fontName,
                        null);
                }
                return;
            }

            if (!FontLog.IsEnabled(logger, FontLogSeverity.Debug)
                && (_resolverExplainsItself || !FontLog.IsEnabled(logger, FontLogSeverity.Information)))
            {
                return;
            }

            var resolvedFamily = FontLog.FamilyOf(font);
            var substituted = resolvedFamily != null
                && !string.Equals(fontName, resolvedFamily, StringComparison.OrdinalIgnoreCase);

            var severity = substituted && !_resolverExplainsItself
                ? FontLogSeverity.Information
                : FontLogSeverity.Debug;

            if (!FontLog.IsEnabled(logger, severity))
                return;

            // For DefaultFontResolver the preceding FontFallback event explains the substitution.
            // A custom resolver does not, so say here that it was the resolver that substituted.
            var message = substituted
                ? string.Format(
                    "Requested font '{0}' {1} was loaded using font '{2}'{3}.",
                    fontName, subFamily, FontLog.Describe(font),
                    _resolverExplainsItself ? string.Empty : " (substituted by the font resolver)")
                : string.Format("Requested font '{0}' {1} was loaded.", fontName, subFamily);

            FontLog.Write(logger, severity, FontLogEventType.FontLoaded, message, fontName, resolvedFamily);
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