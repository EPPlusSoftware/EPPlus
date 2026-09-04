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
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType
{
    /// <summary>
    /// Diagnostic and discovery API over the font file system. Independent of any engine's
    /// configuration: the directories to search are passed in, not read from a configuration,
    /// which is why this never belonged on the engine.
    /// </summary>
    public static class FontDiscovery
    {
        /// <summary>
        /// Returns all available font faces as fully loaded <see cref="OpenTypeFont"/> instances.
        /// Skips corrupt or unreadable fonts but writes diagnostics for each failure.
        /// Not cached, and may take significant time to complete.
        /// </summary>
        public static List<OpenTypeFont> GetAllBaseFontData(
            List<string> fontDirectories,
            bool searchSystemDirectories = true,
            FontFormat? formatTarget = null)
        {
            var locations = DefaultFontLocations.GetLocationsCollection(fontDirectories, searchSystemDirectories);
            var faces = FontScannerV2.EnumerateAllFaces(locations);

            var result = new List<OpenTypeFont>(faces.Count);
            var failures = 0;

            foreach (var face in faces)
            {
                if (formatTarget.HasValue && !MatchesFormat(face.FilePath, formatTarget.Value))
                    continue;

                try
                {
                    var font = new OpenTypeFont(File.ReadAllBytes(face.FilePath));
                    font.EnsureFullyLoaded();
                    result.Add(font);
                }
                catch (Exception ex)
                {
                    failures++;
                    System.Diagnostics.Debug.WriteLine(
                        string.Format("[FontDiscovery] Failed to load font: {0} => {1}: {2}",
                            face.FilePath, ex.GetType().Name, ex.Message));
                }
            }

            if (failures > 0)
            {
                System.Diagnostics.Debug.WriteLine(
                    string.Format("[FontDiscovery] {0} font(s) failed to load.", failures));
            }

            return result;
        }

        private static bool MatchesFormat(string filePath, FontFormat target)
        {
            string ext = Path.GetExtension(filePath);
            if (string.IsNullOrEmpty(ext))
            {
                // No extension: format is undetermined, so do not filter it out.
                return true;
            }

            ext = ext.ToLowerInvariant();
            var format = (ext == ".otf" || ext == ".cff")
                ? FontFormat.Otf
                : FontFormat.Ttf;

            return format == target;
        }
    }
}