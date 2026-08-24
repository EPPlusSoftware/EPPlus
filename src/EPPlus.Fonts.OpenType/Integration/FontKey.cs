using OfficeOpenXml.Interfaces.Fonts;
using System;

namespace EPPlus.Fonts.OpenType.Integration
{
    /// <summary>
    /// Stable identity for a requested font: family name plus EPPlus subfamily
    /// (Regular/Bold/Italic/BoldItalic). This is the single canonical key used to
    /// look up font resources, subset managers and shaped-text providers. It is a
    /// function of the *requested* font only and never depends on the loaded font's
    /// internal name-table values, so it is stable and known before the font is loaded.
    /// </summary>
    public struct FontKey : IEquatable<FontKey>
    {
        public string Family { get; private set; }
        public FontSubFamily SubFamily { get; private set; }

        public FontKey(string family, FontSubFamily subFamily)
        {
            Family = family ?? string.Empty;
            // Normalize the flag combination down to the four canonical values so that
            // any stray flags cannot produce a distinct key for the same logical style.
            SubFamily = Normalize(subFamily);
        }

        private static FontSubFamily Normalize(FontSubFamily subFamily)
        {
            bool bold = (subFamily & FontSubFamily.Bold) == FontSubFamily.Bold;
            bool italic = (subFamily & FontSubFamily.Italic) == FontSubFamily.Italic;
            if (bold && italic) return FontSubFamily.BoldItalic;
            if (bold) return FontSubFamily.Bold;
            if (italic) return FontSubFamily.Italic;
            return FontSubFamily.Regular;
        }

        public bool Equals(FontKey other)
        {
            return SubFamily == other.SubFamily
                && string.Equals(Family, other.Family, StringComparison.Ordinal);
        }

        public override bool Equals(object obj)
        {
            return obj is FontKey && Equals((FontKey)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 31 + (Family != null ? Family.GetHashCode() : 0);
                hash = hash * 31 + (int)SubFamily;
                return hash;
            }
        }

        public override string ToString()
        {
            // Human-readable form, handy for debug output and diagnostics.
            return Family + " " + SubFamily;
        }
    }
}