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
  02/26/2026         EPPlus Software AB           Added overloads accepting prebuilt cache keys
 *************************************************************************************************/
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
        /// Checks if a font is present in the cache (loaded or loading).
        /// </summary>
        public static bool Contains(string cacheKey)
        {
            lock (_syncRoot)
            {
                return _cache.ContainsKey(cacheKey);
            }
        }

        /// <summary>
        /// Creates a placeholder entry to indicate that font loading has begun.
        /// This prevents multiple threads from starting to load the same font.
        /// </summary>
        public static void BeginCache(string cacheKey)
        {
            lock (_syncRoot)
            {
                if (!_cache.ContainsKey(cacheKey))
                {
                    _cache[cacheKey] = new CachedOpenTypeFont()
                    {
                        IsLoaded = false,
                        Font = null
                    };
                }
            }
        }

        /// <summary>
        /// Adds or updates a fully loaded font using a prebuilt cache key.
        /// Signals all waiting threads that the font is now available.
        /// </summary>
        public static void AddToCache(OpenTypeFont font, string cacheKey)
        {
            lock (_syncRoot)
            {
                if (_cache.ContainsKey(cacheKey))
                {
                    _cache[cacheKey].Font = font;
                    _cache[cacheKey].IsLoaded = true;
                }
                else
                {
                    _cache[cacheKey] = new CachedOpenTypeFont
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
        /// Retrieves a font from cache using a prebuilt cache key.
        /// Waits if the font is currently being loaded by another thread.
        /// Returns null if font is not in cache or if timeout occurs while waiting.
        /// </summary>
        public static CachedOpenTypeFont GetFromCache(string cacheKey)
        {
            lock (_syncRoot)
            {
                if (_cache.TryGetValue(cacheKey, out var cached))
                {
                    // If already loaded, return immediately
                    if (cached.IsLoaded && cached.Font != null)
                        return cached;

                    // Wait for another thread to finish loading
                    var timeout = TimeSpan.FromSeconds(2);
                    var start = DateTime.UtcNow;

                    while ((DateTime.UtcNow - start) < timeout)
                    {
                        // CRITICAL: Retrieve from dictionary again after Wait()!
                        // The 'cached' reference may be stale after another thread updates the cache
                        if (_cache.TryGetValue(cacheKey, out cached) && cached.IsLoaded && cached.Font != null)
                            return cached;

                        Monitor.Wait(_syncRoot, TimeSpan.FromMilliseconds(50));
                    }

                    // One final check after timeout
                    if (_cache.TryGetValue(cacheKey, out cached) && cached.IsLoaded && cached.Font != null)
                        return cached;

                    return null;
                }
                return null;
            }
        }
    }
}