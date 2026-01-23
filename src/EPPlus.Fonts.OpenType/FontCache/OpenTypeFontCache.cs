using System;
using System.Collections.Generic;
using System.Threading;

namespace EPPlus.Fonts.OpenType.FontCache
{
    internal static class OpenTypeFontCache
    {
        private static readonly Dictionary<string, CachedOpenTypeFont> _cache =
            new Dictionary<string, CachedOpenTypeFont>(StringComparer.OrdinalIgnoreCase);
        private static readonly object _syncRoot = new object();

        internal static void Clear()
        {
            lock (_syncRoot)
            {
                _cache.Clear();
            }
        }

        /// <summary>
        /// Builds a cache key from family name and subfamily.
        /// THE SUBFAMILY ENUM SHOULD ALWAYS BE INPUT PARAMETER.
        /// Fonts can name themselves however they want, but we map other values in the font to the subfamily.
        /// Therefore never change it to e.g. a string input-parameter
        /// and never use e.g. font.GetEnglishSubfamily name as this could create mismatches.
        /// </summary>
        /// <param name="familyName">Font family name</param>
        /// <param name="subFamily">MUST STAY as FontSubFamily enum</param>
        /// <returns>Cache key string</returns>
        static string BuildCacheKey(string familyName, FontSubFamily subFamily)
        {
            return string.Format("{0}-{1}", familyName, subFamily.ToString());
        }

        /// <summary>
        /// Checks if a font is present in the cache (loaded or loading).
        /// </summary>
        public static bool Contains(string familyName, FontSubFamily subFamily)
        {
            var key = BuildCacheKey(familyName, subFamily);
            lock (_syncRoot)
            {
                return _cache.ContainsKey(key);
            }
        }

        /// <summary>
        /// Creates a placeholder entry to indicate that font loading has begun.
        /// This prevents multiple threads from starting to load the same font.
        /// </summary>
        public static void BeginCache(string familyName, FontSubFamily subFamily)
        {
            lock (_syncRoot)
            {
                var key = BuildCacheKey(familyName, subFamily);
                if (!_cache.ContainsKey(key))
                {
                    _cache[key] = new CachedOpenTypeFont()
                    {
                        IsLoaded = false,
                        Font = null
                    };
                }
            }
        }

        /// <summary>
        /// Adds or updates a fully loaded font in the cache.
        /// Signals all waiting threads that the font is now available.
        /// </summary>
        public static void AddToCache(OpenTypeFont font, string familyName, FontSubFamily subFamily)
        {
            lock (_syncRoot)
            {
                var key = BuildCacheKey(familyName, subFamily);

                // Update existing entry OR create new one
                if (_cache.ContainsKey(key))
                {
                    _cache[key].Font = font;
                    _cache[key].IsLoaded = true;
                }
                else
                {
                    _cache[key] = new CachedOpenTypeFont
                    {
                        Font = font,
                        IsLoaded = true
                    };
                }

                // Signal all waiting threads that the font is loaded
                Monitor.PulseAll(_syncRoot);
            }
        }

        /// <summary>
        /// Retrieves a font from cache, waiting if it's currently being loaded by another thread.
        /// Returns null if font is not in cache or if timeout occurs while waiting.
        /// </summary>
        /// <param name="familyName">Font family name</param>
        /// <param name="subFamily">Font subfamily</param>
        /// <returns>Cached font entry or null if not available</returns>
        public static CachedOpenTypeFont GetFromCache(string familyName, FontSubFamily subFamily)
        {
            var key = BuildCacheKey(familyName, subFamily);
            lock (_syncRoot)
            {
                if (_cache.TryGetValue(key, out var cached))
                {
                    // If already loaded, return immediately
                    if (cached.IsLoaded && cached.Font != null)
                    {
                        return cached;
                    }

                    // Wait for another thread to finish loading
                    var timeout = TimeSpan.FromSeconds(2);
                    var start = DateTime.UtcNow;

                    while ((DateTime.UtcNow - start) < timeout)
                    {
                        // CRITICAL: Retrieve from dictionary again after Wait()!
                        // The 'cached' reference may be stale after another thread updates the cache
                        if (_cache.TryGetValue(key, out cached) && cached.IsLoaded && cached.Font != null)
                        {
                            return cached;
                        }

                        // Wait and release lock temporarily
                        Monitor.Wait(_syncRoot, TimeSpan.FromMilliseconds(50));
                    }

                    // Timeout occurred - one final check
                    if (_cache.TryGetValue(key, out cached) && cached.IsLoaded && cached.Font != null)
                    {
                        return cached;
                    }

                    // Timeout without result
                    return null;
                }
                return null;
            }
        }
    }
}