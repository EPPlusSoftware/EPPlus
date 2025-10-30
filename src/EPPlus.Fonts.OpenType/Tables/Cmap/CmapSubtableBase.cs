using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public abstract class CmapSubtableBase : FontTableElement
    {
        /// <summary>
        /// Format identifier (0, 4, 6, etc.)
        /// </summary>
        public abstract ushort Format { get; }

        /// <summary>
        /// Length of the subtable in bytes
        /// </summary>
        public abstract ushort Length { get; }

        /// <summary>
        /// Language code (optional usage depending on format)
        /// </summary>
        public abstract ushort Language { get; }

        /// <summary>
        /// Array of glyph mappings (character code → glyph index)
        /// </summary>
        public abstract GlyphMapping[] GlyphMappingArray { get; }


    }
}
