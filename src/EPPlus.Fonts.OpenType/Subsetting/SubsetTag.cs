using EPPlus.Fonts.OpenType.Integration;
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal static class SubsetTag
    {
        /// <summary>
        /// Creates a deterministic PDF subset tag ("ABCDEF+") from the font identity
        /// and the code points included in the subset.
        /// </summary>
        internal static string Create(FontKey identity, IEnumerable<int> codePoints)
        {
            unchecked
            {
                uint hash = 2166136261;
                foreach (var c in identity.Family ?? string.Empty)
                    hash = (hash ^ c) * 16777619;
                hash = (hash ^ (uint)(int)identity.SubFamily) * 16777619;

                if (codePoints != null)
                {
                    var sorted = new List<int>(codePoints);
                    sorted.Sort();
                    foreach (var cp in sorted)
                        hash = (hash ^ (uint)cp) * 16777619;
                }

                // 26^6 = 308 915 776 < 2^32
                uint value = hash % 308915776u;
                var chars = new char[6];
                for (int i = 5; i >= 0; i--)
                {
                    chars[i] = (char)('A' + (int)(value % 26));
                    value /= 26;
                }
                return new string(chars) + "+";
            }
        }
    }
}