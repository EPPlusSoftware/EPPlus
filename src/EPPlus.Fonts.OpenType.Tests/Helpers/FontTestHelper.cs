using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.Helpers
{
    /// <summary>
    /// Shared test utilities for font testing
    /// </summary>
    public static class FontTestHelper
    {
        /// <summary>
        /// Counts the total number of ligatures in a font's GSUB table
        /// </summary>
        /// <param name="font">Font to analyze</param>
        /// <returns>Total number of ligature rules</returns>
        public static int CountLigatures(OpenTypeFont font)
        {
            if (font?.GsubTable == null)
                return 0;

            int count = 0;

            foreach (var lookup in font.GsubTable.LookupList.Lookups)
            {
                // Lookup Type 4 = Ligature Substitution
                if (lookup.LookupType == 4)
                {
                    foreach (var subtable in lookup.SubTables)
                    {
                        var ligSubtable = subtable as LigatureSubstSubTable;
                        if (ligSubtable?.LigatureSets != null)
                        {
                            foreach (var ligatureSet in ligSubtable.LigatureSets.Values)
                            {
                                count += ligatureSet.Ligatures.Count;
                            }
                        }
                    }
                }
            }

            return count;
        }

        /// <summary>
        /// Asserts that a font passes validation
        /// </summary>
        /// <param name="font">Font to validate</param>
        /// <param name="severity">Minimum severity level for errors</param>
        public static void AssertFontValid(
            OpenTypeFont font,
            FontValidationSeverity severity = FontValidationSeverity.Error)
        {
            var validator = new FontValidator();
            var report = validator.Validate(font, severity);

            if (!report.IsValid)
            {
                var errorMessages = report.Errors
                    .Select(e => $"  [{e.Severity}] {e.ParentResult.TableName ?? "General"}: {e.Message}")
                    .ToList();

                var message = string.Format(
                    "Font validation failed with {0} error(s):\n{1}",
                    report.Errors.Count(),
                    string.Join("\n", errorMessages));

                throw new AssertFailedException(message);
            }
        }

        /// <summary>
        /// Creates a subset and serializes it to bytes
        /// </summary>
        /// <param name="fontName">Name of font to load</param>
        /// <param name="text">Text to subset</param>
        /// <param name="fontFolders">Folders to search for fonts</param>
        /// <returns>Serialized subset bytes</returns>
        public static byte[] SubsetAndSerialize(
            string fontName,
            string text,
            List<string> fontFolders)
        {
            var font = OpenTypeFonts.LoadFont(fontName, FontSubFamily.Regular);
            var subset = font.CreateSubset(text);
            return subset.Serialize();
        }

        /// <summary>
        /// Creates a subset and serializes it to bytes (char array overload)
        /// </summary>
        public static byte[] SubsetAndSerialize(
            string fontName,
            char[] chars,
            List<string> fontFolders)
        {
            var font = OpenTypeFonts.LoadFont(fontName, FontSubFamily.Regular);
            var subset = font.CreateSubset(chars);
            return subset.Serialize();
        }

        /// <summary>
        /// Performs a full roundtrip: subset → serialize → parse → validate
        /// </summary>
        /// <param name="fontName">Name of font to load</param>
        /// <param name="text">Text to subset</param>
        /// <param name="fontFolders">Folders to search for fonts</param>
        /// <returns>Parsed subset font (validated)</returns>
        public static OpenTypeFont RoundtripSubset(
            string fontName,
            string text,
            List<string> fontFolders)
        {
            var font = OpenTypeFonts.LoadFont(fontName, FontSubFamily.Regular);
            var subset = font.CreateSubset(text);
            var bytes = subset.Serialize();

            var parsed = new OpenTypeFont(bytes);

            AssertFontValid(parsed);

            return parsed;
        }

        /// <summary>
        /// Performs a full roundtrip with char array
        /// </summary>
        public static OpenTypeFont RoundtripSubset(
            string fontName,
            char[] chars,
            List<string> fontFolders)
        {
            var font = OpenTypeFonts.LoadFont(fontName, FontSubFamily.Regular);
            var subset = font.CreateSubset(chars);
            var bytes = subset.Serialize();

            var parsed = new OpenTypeFont(bytes);

            AssertFontValid(parsed);

            return parsed;
        }

        /// <summary>
        /// Gets the total number of glyphs in a subset for specific characters
        /// </summary>
        /// <param name="font">Original font</param>
        /// <param name="chars">Characters to check</param>
        /// <returns>Number of unique glyph IDs needed</returns>
        public static int GetExpectedGlyphCount(OpenTypeFont font, char[] chars)
        {
            var glyphIds = new HashSet<ushort>();

            // Always include .notdef
            glyphIds.Add(0);

            // Map characters to glyph IDs
            foreach (var ch in chars)
            {
                if (font.CmapTable.TryGetGlyphId(ch, out ushort glyphId))
                {
                    glyphIds.Add(glyphId);
                }
            }

            return glyphIds.Count;
        }

        /// <summary>
        /// Checks if a font has a specific lookup type in GSUB
        /// </summary>
        public static bool HasGsubLookupType(OpenTypeFont font, ushort lookupType)
        {
            if (font?.GsubTable == null)
                return false;

            return font.GsubTable.LookupList.Lookups
                .Any(lookup => lookup.LookupType == lookupType);
        }

        /// <summary>
        /// Gets all ligature lookup types present in font
        /// </summary>
        public static List<ushort> GetGsubLookupTypes(OpenTypeFont font)
        {
            if (font?.GsubTable == null)
                return new List<ushort>();

            return font.GsubTable.LookupList.Lookups
                .Select(lookup => lookup.LookupType)
                .Distinct()
                .OrderBy(type => type)
                .ToList();
        }

        /// <summary>
        /// Saves a font to a file (useful for manual inspection)
        /// </summary>
        public static void SaveFontForInspection(
            OpenTypeFont font,
            string filename)
        {
            var bytes = font.Serialize();
            var tempPath = Path.Combine(Path.GetTempPath(), filename);
            File.WriteAllBytes(tempPath, bytes);
            Console.WriteLine($"Font saved to: {tempPath}");
        }
    }
}