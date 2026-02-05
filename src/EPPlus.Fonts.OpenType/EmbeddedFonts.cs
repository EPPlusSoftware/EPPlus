using EPPlus.Fonts.OpenType.Scanner;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Loads embedded fallback fonts.
    /// Internal - not exposed to users.
    /// </summary>
    internal static class EmbeddedFonts
    {
        private static readonly Dictionary<string, OpenTypeFont> _cache =
            new Dictionary<string, OpenTypeFont>();

        private static readonly object _lock = new object();

        /// <summary>
        /// Loads Noto Emoji Regular (embedded resource).
        /// Cached after first load.
        /// </summary>
        internal static OpenTypeFont LoadNotoEmoji()
        {
            return LoadCached("NotoEmoji-Regular.ttf");
        }

        private static OpenTypeFont LoadCached(string resourceName)
        {
            lock (_lock)
            {
                if (_cache.TryGetValue(resourceName, out var font))
                    return font;

                var assembly = Assembly.GetExecutingAssembly();
                var fullResourceName = $"EPPlus.Fonts.OpenType.Resources.{resourceName}";

                using (var stream = assembly.GetManifestResourceStream(fullResourceName))
                {
                    if (stream == null)
                    {
                        throw new InvalidOperationException(
                            $"Embedded font resource not found: {resourceName}. " +
                            "This is a bug in EPPlus.Fonts.OpenType - please report it.");
                    }

                    font = OpenTypeFonts.GetFromBytes(bytes: ReadStreamFully(stream), FontFormat.Ttf);
                    _cache[resourceName] = font;
                    return font;
                }
            }
        }

        private static byte[] ReadStreamFully(Stream stream)
        {
            throw new NotImplementedException();
        }
    }
}