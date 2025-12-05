using System.Collections.Generic;
using System.IO;
using System;
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
            var directories = OpenTypeFonts.GetLocationsCollection(additionalDirectories, searchSystemDirectories);
            var candidates = EnumerateAllFaces(directories);

            FontFaceInfo bestMatch = null;
            int bestScore = -1;

            foreach (var face in candidates)
            {
                if (string.IsNullOrEmpty(face.FamilyName))
                    continue;

                int score = CalculateMatchScore(face, familyName, desiredStyle);

                if (score > bestScore)
                {
                    bestScore = score;
                    bestMatch = face;
                }
            }

            return bestMatch;
        }

        private static int CalculateMatchScore(FontFaceInfo face, string requestedFamily, FontSubFamily requestedStyle)
        {
            int score = 0;

            string faceFamilyLower = (face.FamilyName ?? "").ToLowerInvariant();
            string requestedLower = requestedFamily.ToLowerInvariant();

            // Exact family name → decisive win
            if (string.Equals(face.FamilyName, requestedFamily, StringComparison.OrdinalIgnoreCase))
                score += 10_000;

            // One name is substring of the other (e.g. "Aptos Narrow" vs "Aptos")
            else if (faceFamilyLower.Contains(requestedLower) || requestedLower.Contains(faceFamilyLower))
                score += 5_000;

            // Partial overlap
            else if (faceFamilyLower.IndexOf(requestedLower, StringComparison.OrdinalIgnoreCase) >= 0 ||
                     requestedLower.IndexOf(faceFamilyLower, StringComparison.OrdinalIgnoreCase) >= 0)
                score += 1_000;

            // Style matching
            if (face.Subfamily == requestedStyle)
                score += 2_000;
            else if (requestedStyle == FontSubFamily.Regular || face.Subfamily == FontSubFamily.Regular)
                score += 500;                       // Regular is acceptable fallback
            else if ((requestedStyle & face.Subfamily) != 0) // BoldItalic contains Bold, etc.
                score += 1_000;

            return score;
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