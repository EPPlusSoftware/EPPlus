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
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Thread-safe cache for scanned font faces.
    /// Key = FilePath|OffsetInFile, Value = FontFaceInfo
    /// Automatically invalidates entries if file has been modified.
    /// </summary>
    internal static class FontScannerCache
    {
        private static readonly Dictionary<string, FontFaceInfo> _cache
            = new Dictionary<string, FontFaceInfo>(StringComparer.OrdinalIgnoreCase);

        private static readonly object _syncLock = new object();

        /// <summary>
        /// Returns cached FontFaceInfo or creates and caches a new one using the provided factory.
        /// </summary>
        public static FontFaceInfo GetOrAdd(string filePath, long offset, Func<string, long, FontFaceInfo> factory)
        {
            string key = filePath + "|" + offset;

            lock (_syncLock)
            {
                if (_cache.TryGetValue(key, out FontFaceInfo info))
                {
                    // Invalidate if file was modified
                    try
                    {
                        DateTime current = File.GetLastWriteTimeUtc(filePath);
                        if (current != info.LastWriteTimeUtc)
                        {
                            _cache.Remove(key);
                        }
                        else
                        {
                            return info;
                        }
                    }
                    catch
                    {
                        // File deleted or access denied → remove from cache
                        _cache.Remove(key);
                    }
                }

                info = factory(filePath, offset);
                _cache[key] = info;
                return info;
            }
        }

        /// <summary>
        /// Clears the entire cache. Used during testing and when font folders change.
        /// </summary>
        public static void Clear()
        {
            lock (_syncLock)
            {
                _cache.Clear();
            }
        }

        /// <summary>
        /// Returns all cached faces (for diagnostics or full enumeration).
        /// </summary>
        public static List<FontFaceInfo> GetAll()
        {
            lock (_syncLock)
            {
                return new List<FontFaceInfo>(_cache.Values);
            }
        }
    }
}
