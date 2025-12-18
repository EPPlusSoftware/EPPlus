using System;
using System.Collections.Generic;
using System.IO;
using System.Security;

namespace EPPlus.Fonts.OpenType.Scanner
{
    /// <summary>
    /// Thread-safe cache for directory listings of font files.
    /// Prevents repeated (and expensive) Directory.GetFiles() calls on the same font directories.
    /// Especially important for system font folders (C:\Windows\Fonts) with thousands of files.
    /// Fully compatible with .NET 3.5.
    /// </summary>
    internal static class FontDirectoryCache
    {
        private static readonly Dictionary<string, CachedDirectoryInfo> _cache
            = new Dictionary<string, CachedDirectoryInfo>(StringComparer.OrdinalIgnoreCase);

        private static readonly object _lock = new object();

        private class CachedDirectoryInfo
        {
            public string[] FontFiles;       // List of .ttf, .otf, .ttc files
            public DateTime LastCheckTime;   // When we last validated the cache
            public long DirectoryTicks;      // Directory.GetLastWriteTimeUtc().Ticks
        }

        /// <summary>
        /// Returns cached list of font files in the given directory.
        /// If the directory has not changed since last check, returns cached result.
        /// Otherwise performs a fresh scan and updates the cache.
        /// </summary>
        public static string[] GetFontFiles(string directory)
        {
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                return new string[0];

            lock (_lock)
            {
                CachedDirectoryInfo cached;
                bool needRefresh = true;

                if (_cache.TryGetValue(directory, out cached))
                {
                    try
                    {
                        // Simple but effective: compare directory write time
                        long currentTicks = Directory.GetLastWriteTimeUtc(directory).Ticks;
                        if (currentTicks == cached.DirectoryTicks)
                            needRefresh = false;
                    }
                    catch (UnauthorizedAccessException) { /* ignore */ }
                    catch (IOException) { /* ignore */ }
                    catch (SecurityException) { /* ignore */ }
                }

                if (!needRefresh)
                    return cached.FontFiles;

                // Perform fresh scan
                string[] files = ScanDirectoryForFonts(directory);

                // Update cache
                cached = new CachedDirectoryInfo
                {
                    FontFiles = files,
                    LastCheckTime = DateTime.UtcNow,
                    DirectoryTicks = Directory.GetLastWriteTimeUtc(directory).Ticks
                };

                _cache[directory] = cached;
                return files;
            }
        }

        /// <summary>
        /// Performs the actual directory scan for font files.
        /// Separated for clarity and testability.
        /// </summary>
        private static string[] ScanDirectoryForFonts(string directory)
        {
            try
            {
                string[] allFiles = Directory.GetFiles(directory, "*.*", SearchOption.AllDirectories);
                List<string> fontFiles = new List<string>(allFiles.Length);

                foreach (string file in allFiles)
                {
                    string ext = Path.GetExtension(file);
                    if (ext != null)
                    {
                        ext = ext.ToLowerInvariant();
                        if (ext == ".ttf" || ext == ".otf" || ext == ".ttc")
                        {
                            fontFiles.Add(file);
                        }
                    }
                }

                return fontFiles.ToArray();
            }
            catch (UnauthorizedAccessException) { return new string[0]; }
            catch (IOException) { return new string[0]; }
            catch (SecurityException) { return new string[0]; }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("[FontDirectoryCache] Failed to scan directory: " + directory + " → " + ex.Message);
                return new string[0];
            }
        }

        /// <summary>
        /// Clears the entire directory cache.
        /// Call when font folders may have changed or during application reset.
        /// </summary>
        public static void Clear()
        {
            lock (_lock)
            {
                _cache.Clear();
            }
        }
    }
}