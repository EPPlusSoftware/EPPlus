/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2026         EPPlus Software AB           OpenType font implementation
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Scanner
{
    internal static partial class FontScannerV2
    {
        public static FontFaceInfo FindBestMatch(
            string additionalDirectory,
            string familyName,
            FontSubFamily desiredStyle,
            bool searchSystemDirectories = true)
        {
            var dirs = new List<string>() { additionalDirectory };
            return FindBestMatch(dirs, familyName, desiredStyle, searchSystemDirectories);
        }

        public static FontFaceInfo FindBestMatch(
            IEnumerable<string> additionalDirectories,
            string familyName,
            FontSubFamily desiredStyle,
            bool searchSystemDirectories = true)
        {
            var directories = DefaultFontLocations.GetLocationsCollection(additionalDirectories, searchSystemDirectories);
            var candidates = EnumerateAllFaces(directories);

            FontFaceInfo bestMatch = null;
            int bestScore = -1;
            int nMatches = 0;
            foreach (var face in candidates)
            {
                if (string.IsNullOrEmpty(face.FamilyName))
                    continue;

                int score = CalculateMatchScore(face, familyName, desiredStyle);

                
                if (score > bestScore)
                {
                    if (score >= 3000) nMatches++;
                    bestScore = score;
                    bestMatch = face;
                }
            }
            if (bestMatch == null)
                return null;

            // Don't mutate the cached face. The cache returns the same FontFaceInfo instance to all
            // callers, and IsExactMatch is per-query state, not a property of the font on disk.
            // Mutating the cached instance creates a race condition between parallel callers.
            var result = bestMatch.Clone();
            result.IsExactMatch = bestScore >= 9_000;
            return result;
        }

        private static int CalculateMatchScore(FontFaceInfo face, string requestedFamily, FontSubFamily requestedStyle)
        {
            int score = 0;

            // Normalize: remove whitespace and convert to lowercase for comparison
            string faceFamily = face.FamilyName ?? "";
            string faceFamilyNormalized = NormalizeFontName(faceFamily);
            string requestedNormalized = NormalizeFontName(requestedFamily);

            // Exact family name (case-insensitive) → decisive win
            if (string.Equals(faceFamily, requestedFamily, StringComparison.OrdinalIgnoreCase))
                score += 10_000;
            // Exact match after normalization (whitespace removed)
            else if (faceFamilyNormalized == requestedNormalized)
                score += 9_000;
            // One name is substring of the other (e.g. "Aptos Narrow" vs "Aptos")
            else if (faceFamilyNormalized.Contains(requestedNormalized) ||
                     requestedNormalized.Contains(faceFamilyNormalized))
                score += 5_000;
            // Partial overlap - using IndexOf with StringComparison
            else if (faceFamilyNormalized.IndexOf(requestedNormalized, StringComparison.Ordinal) >= 0 ||
                     requestedNormalized.IndexOf(faceFamilyNormalized, StringComparison.Ordinal) >= 0)
                score += 1_000;

            // Style matching
            if (face.Subfamily == requestedStyle)
                score += 2_000;
            else if (requestedStyle == FontSubFamily.Regular || face.Subfamily == FontSubFamily.Regular)
                score += 500;
            else if ((requestedStyle & face.Subfamily) != 0)
                score += 1_000;

            return score;
        }

        /// <summary>
        /// Normalizes a font name for fuzzy matching.
        /// Removes whitespace, hyphens, and converts to lowercase.
        /// </summary>
        private static string NormalizeFontName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return string.Empty;

            // Remove common separators
            return name.Replace(" ", "")
                       .Replace("-", "")
                       .Replace("_", "")
                       .ToLowerInvariant();
        }

        internal static List<FontFaceInfo> EnumerateAllFaces(List<string> directories)
        {
            var result = new List<FontFaceInfo>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string dir in directories)
            {
                if (!Directory.Exists(dir)) continue;

                string[] files = FontDirectoryCache.GetFontFiles(dir);

                foreach (string file in files)
                {
                    if (!seen.Add(file)) continue;

                    try
                    {
                        if (Path.GetExtension(file).Equals(".ttc", StringComparison.OrdinalIgnoreCase))
                        {
                            result.AddRange(ScanTtcFile(file));
                        }
                        else
                        {
                            var face = FontScannerCache.GetOrAdd(file, 0, FontScannerV2Core.ScanSingleFace);
                            result.Add(face);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[FontScannerV2] Failed to read font: {file} → {ex.GetType().Name}: {ex.Message}");
                    }
                }
            }

            return result;
        }

        internal static FontFaceInfo GetFace(string filePath, long offset = 0)
        {
            return FontScannerCache.GetOrAdd(filePath, offset, FontScannerV2Core.ScanSingleFace);
        }

        /// <summary>
        /// Returns all scanned font faces from a specific directory (and subdirectories).
        /// Uses the same high-performance, cached scanning as FindBestMatch.
        /// Perfect for diagnostics, font picker UI, or when you need to list all available fonts in a folder.
        /// </summary>
        /// <param name="path">The directory to scan (e.g. @"C:\Windows\Fonts")</param>
        /// <returns>List of FontFaceInfo for all valid TTF/OTF/TTC faces found</returns>
        internal static List<FontFaceInfo> GetAllScannedFontsInPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return new List<FontFaceInfo>();

            if (!Directory.Exists(path))
                return new List<FontFaceInfo>();

            var directories = new List<string> { path };
            return EnumerateAllFaces(directories);
        }
    }
}