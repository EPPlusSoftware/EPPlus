/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/26/2021         EPPlus Software AB       EPPlus 6.0
  09/01/2026         EPPlus Software AB       Per-font loading
 *************************************************************************************************/
using OfficeOpenXml.Packaging.Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace EPPlus.Fonts.OpenType.GenericFontWidths
{
    /// <summary>
    /// Loads serialized font metrics from the resources/TextMetrics.zip archive.
    /// </summary>
    internal static class GenericFontMetricsLoader
    {
        private const string ResourceName = "EPPlus.Fonts.OpenType.Resources.TextMetrics.zip";

        /// <summary>
        /// Loads every font in the archive.
        /// </summary>
        internal static Dictionary<uint, SerializedFontMetrics> LoadFontMetrics()
        {
            var fonts = new Dictionary<uint, SerializedFontMetrics>();
            using (var stream = GetResourceStream())
            {
                var zipStream = new ZipInputStream(stream);
                ZipEntry entry;
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (entry.IsDirectory || Path.GetExtension(entry.FileName) != ".fmtr") continue;

                    var metrics = ReadEntry(zipStream, entry);
                    fonts[metrics.GetKey()] = metrics;
                }
            }
            return fonts;
        }

        /// <summary>
        /// The font keys present in the archive, taken from the file names without reading or
        /// decompressing any of them. Lets the caller answer "does this font exist" without
        /// paying for the metrics of fonts nobody asked about.
        /// </summary>
        internal static HashSet<uint> LoadAvailableFontKeys()
        {
            var keys = new HashSet<uint>();
            using (var stream = GetResourceStream())
            {
                var zipStream = new ZipInputStream(stream);
                ZipEntry entry;
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (entry.IsDirectory || Path.GetExtension(entry.FileName) != ".fmtr") continue;

                    uint key;
                    if (TryGetKeyFromFileName(entry.FileName, out key))
                    {
                        keys.Add(key);
                    }
                }
            }
            return keys;
        }

        /// <summary>
        /// Loads a single font. Returns null when the archive holds no such font.
        ///
        /// Most workbooks use one or two fonts, so loading all 101 on first measurement is
        /// almost entirely wasted. The file name is the font key, so the right entry can be
        /// found without decompressing the others.
        /// </summary>
        internal static SerializedFontMetrics LoadFontMetrics(uint fontKey)
        {
            using (var stream = GetResourceStream())
            {
                var zipStream = new ZipInputStream(stream);
                ZipEntry entry;
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (entry.IsDirectory || Path.GetExtension(entry.FileName) != ".fmtr") continue;

                    uint key;
                    if (!TryGetKeyFromFileName(entry.FileName, out key) || key != fontKey) continue;

                    return ReadEntry(zipStream, entry);
                }
            }
            return null;
        }

        private static Stream GetResourceStream()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var stream = assembly.GetManifestResourceStream(ResourceName);
            if (stream == null)
            {
                throw new InvalidOperationException("Embedded resource not found: " + ResourceName);
            }
            return stream;
        }

        private static bool TryGetKeyFromFileName(string fileName, out uint key)
        {
            var name = Path.GetFileNameWithoutExtension(fileName);
            return uint.TryParse(name, out key);
        }

        private static SerializedFontMetrics ReadEntry(ZipInputStream zipStream, ZipEntry entry)
        {
            var bytes = ReadExactly(zipStream, (int)entry.UncompressedSize);
            using (var ms = new MemoryStream(bytes))
            {
                return GenericFontMetricsSerializer.Deserialize(ms);
            }
        }

        /// <summary>
        /// Reads exactly count bytes.
        ///
        /// The previous code called Read once and assigned the result to an unused variable. A
        /// decompressing stream is allowed to return fewer bytes than asked for, and when it
        /// did the shortfall showed up as a truncated font file rather than an error.
        /// </summary>
        private static byte[] ReadExactly(Stream stream, int count)
        {
            var buffer = new byte[count];
            var offset = 0;
            while (offset < count)
            {
                var read = stream.Read(buffer, offset, count - offset);
                if (read <= 0)
                {
                    throw new EndOfStreamException(
                        string.Format("Expected {0} bytes of font metrics but the stream ended after {1}.",
                                      count, offset));
                }
                offset += read;
            }
            return buffer;
        }
    }
}