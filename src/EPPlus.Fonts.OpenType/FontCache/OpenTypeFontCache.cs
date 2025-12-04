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

        private static readonly object _logLock = new object();
        private const string LogFilePath = @"c:\Temp\cachelog.txt";


        internal static void LogAccess(string message)
        {
            lock (_logLock)
            {
                File.AppendAllText(LogFilePath, $"{DateTime.UtcNow:O} - {message}{Environment.NewLine}");
            }
        }

        internal static void Clear()
        {
            lock (_syncRoot)
            {
                _cache.Clear();
            }
        }



        static string BuildCacheKey(string familyName, string subFamily)
        {
            return $"{familyName}-{subFamily}";
        }

        public static bool Contains(string familyName, FontSubFamily subFamily)
        {
            var key = BuildCacheKey(familyName, subFamily.ToString());
            lock (_syncRoot)
            {

                bool exists = _cache.ContainsKey(key);
                LogAccess($"Contains called for key '{key}' -> {exists}");
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
                    LogAccess($"BeginCache: Added placeholder for key '{key}'");
                }
                else
                {
                    LogAccess($"BeginCache: Key '{key}' already exists");
                }
            }
        }

        public static void AddToCache(OpenTypeFont font)
        {
            lock (_syncRoot)
            {
                var key = BuildCacheKey(font.GetEnglishFullFontFamilyName(), font.GetEnglishFontSubFamilyName());
                if (!_cache.ContainsKey(key))
                {
                    _cache[key] = new CachedOpenTypeFont();
                    LogAccess($"AddToCache: Created new entry for key '{key}', full name: {font.FullName}");
                }
                else
                {
                    LogAccess($"AddToCache: Updated existing entry for key '{key}', full name: {font.FullName}");
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
                        LogAccess($"GetFromCache: Key '{key}' returned immediately. Full name: {cached.Font.FullName}");
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
                        LogAccess($"GetFromCache: Key '{key}' not found after wait.");
                        return null;
                    }
                    LogAccess($"GetFromCache: Key '{key}' returned after wait -> Loaded={cached.IsLoaded}, Full name: {cached.Font.FullName}");
                    return cached.IsLoaded ? cached : null;
                }

                return null;
            }
        }
    }

}
