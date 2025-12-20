using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups
{
    /// <summary>
    /// En hjälpstruktur för att hålla mappningen mellan gamla och nya Glyph IDs 
    /// under omskrivningen (Rewrite) av GSUB-tabeller.
    /// </summary>
    internal struct GsubRewriteEntry
    {
        /// <summary>
        /// Det nya Glyph ID:t för källtecknet (Input).
        /// </summary>
        public ushort NewInput;

        /// <summary>
        /// Det nya Glyph ID:t för ersättningstecknet (Output/Substitute).
        /// </summary>
        public ushort NewOutput;
    }
}