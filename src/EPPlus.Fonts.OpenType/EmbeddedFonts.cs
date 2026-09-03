using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Fonts;
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

        internal static OpenTypeFont LoadNotoMath()
        {
            return LoadCached("NotoSansMath-Regular.ttf");
        }

        internal static OpenTypeFont LoadArchivoNarrow(FontSubFamily subFamily)
        {
            if (subFamily == FontSubFamily.Italic)
            {
                return LoadCached("ArchivoNarrow-Italic.ttf");
            }
            else if (subFamily == FontSubFamily.Bold)
            {
                return LoadCached("ArchivoNarrow-Bold.ttf");
            }
            else if(subFamily == FontSubFamily.BoldItalic)
            {
                return LoadCached("ArchivoNarrow-BoldItalic.ttf");
            }
            return LoadCached("ArchivoNarrow-Regular.ttf");
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

                    font = new OpenTypeFont(fontBytes: ReadStreamFully(stream));
                    font.EnsureFullyLoaded();
                    _cache[resourceName] = font;
                    return font;
                }
            }
        }

        /// <summary>
        /// Reads all bytes from a stream into a byte array.
        /// .NET 3.5 compatible (no CopyTo available).
        /// </summary>
        private static byte[] ReadStreamFully(Stream stream)
        {
            if (stream == null)
                throw new ArgumentNullException("stream");

            // Try to use stream length if available (e.g., MemoryStream, FileStream)
            if (stream.CanSeek)
            {
                long length = stream.Length - stream.Position;
                if (length > int.MaxValue)
                    throw new InvalidOperationException("Stream is too large to read into memory");

                byte[] buffer = new byte[(int)length];
                int offset = 0;
                int remaining = (int)length;

                while (remaining > 0)
                {
                    int read = stream.Read(buffer, offset, remaining);
                    if (read <= 0)
                        throw new EndOfStreamException("Stream ended before reading all bytes");

                    offset += read;
                    remaining -= read;
                }

                return buffer;
            }
            else
            {
                // Non-seekable stream (rare for embedded resources, but handle it)
                // Use MemoryStream to accumulate bytes
                using (var ms = new MemoryStream())
                {
                    byte[] buffer = new byte[8192]; // 8KB buffer
                    int bytesRead;

                    while ((bytesRead = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        ms.Write(buffer, 0, bytesRead);
                    }

                    return ms.ToArray();
                }
            }
        }
    }
}