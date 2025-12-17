using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace EPPlus.Fonts.OpenType.FontCache
{

    internal static class OpenTypeFontCache
    {
        private static readonly Dictionary<string, CachedOpenTypeFont> _cache = new Dictionary<string, CachedOpenTypeFont>(StringComparer.OrdinalIgnoreCase);
        private static readonly object _syncRoot = new object();

        internal static void Clear()
        {
            lock (_syncRoot)
            {
                _cache.Clear();
            }
        }



        internal static string BuildCacheKey(string familyName, string subFamily)
        {
            return $"{familyName}-{subFamily}";
        }

        public static bool Contains(string familyName, FontSubFamily subFamily)
        {
            var key = BuildCacheKey(familyName, subFamily.ToString());
            lock (_syncRoot)
            {

                bool exists = _cache.ContainsKey(key);
                return exists;

            }
        }

        public static void BeginCache(string familyName, FontSubFamily subFamily)
        {
            lock (_syncRoot)
            {
                var key = BuildCacheKey(familyName, subFamily.ToString());
                if (!_cache.ContainsKey(key))
                {
                    _cache[key] = new CachedOpenTypeFont()
                    {
                        IsLoaded = false
                    };
                }
            }
        }

        public static void AddToCache(OpenTypeFont font, string key)
        {
            lock (_syncRoot)
            {
                //var key = BuildCacheKey(font.GetEnglishFontFamilyName(), font.GetEnglishFontSubFamilyName());
                if (!_cache.ContainsKey(key))
                {
                    _cache[key] = new CachedOpenTypeFont();
                }


                _cache[key].Font = font;
                _cache[key].IsLoaded = true;

                // Signalera alla väntande trådar att fonten är laddad
                Monitor.PulseAll(_syncRoot);
            }
        }

        public static CachedOpenTypeFont GetFromCache(string familyName, FontSubFamily subFamily)
        {
            var key = BuildCacheKey(familyName, subFamily.ToString());
            lock (_syncRoot)
            {
                if (_cache.TryGetValue(key, out var cached))
                {
                    if (cached.IsLoaded)
                    {
                        return cached;
                    }

                    // Wait max 1 second for the font to be loaded
                    var timeout = TimeSpan.FromSeconds(2);
                    var start = DateTime.UtcNow;

                    while (!cached.IsLoaded && (DateTime.UtcNow - start) < timeout)
                    {
                        Monitor.Wait(_syncRoot, TimeSpan.FromMilliseconds(50));
                    }
                    if (cached == null || cached.Font == null)
                    {
                        return null;
                    }
                    return cached.IsLoaded ? cached : null;
                }

                return null;
            }
        }
    }
}
