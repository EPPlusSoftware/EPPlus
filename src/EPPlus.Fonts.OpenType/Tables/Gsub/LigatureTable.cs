using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Gsub
{
    public class LigatureTable : FontTableElement
    {
        /// <summary>
        /// The Glyph ID of the resulting ligature glyph (output).
        /// </summary>
        public ushort LigatureGlyph { get; set; }

        /// <summary>
        /// The array of component Glyph IDs that follow the initial glyph (input sequence).
        /// Note: The initial glyph is implicitly defined by the Coverage table that points to the LigatureSet.
        /// </summary>
        public ushort[] Components { get; set; }


        /// <summary>
        /// Remaps the output ligature glyph ID and all component glyph IDs 
        /// from old IDs to new subset IDs.
        /// </summary>
        /// <param name="oldToNewGlyphId">The mapping dictionary.</param>
        /// <returns>True if the LigatureGlyph is included in the subset, otherwise false.</returns>
        internal bool Remap(Dictionary<ushort, ushort> oldToNewGlyphId)
        {
            ushort newLigatureGid;

            // 1. Remap the resulting Ligature Glyph ID
            if (oldToNewGlyphId.TryGetValue(this.LigatureGlyph, out newLigatureGid))
            {
                this.LigatureGlyph = newLigatureGid;

                // 2. Remap component glyph IDs (input sequence)
                for (int i = 0; i < this.Components.Length; i++)
                {
                    ushort oldComponentGid = this.Components[i];
                    ushort newComponentGid;

                    if (oldToNewGlyphId.TryGetValue(oldComponentGid, out newComponentGid))
                    {
                        this.Components[i] = newComponentGid;
                    }
                    else
                    {
                        // If a component is not in the subset, the ligature is invalid 
                        // and should be discarded later. For now, remap to .notdef (New ID 0).
                        this.Components[i] = 0;
                    }
                }
                return true;
            }

            // Ligature output glyph is not in the subset, so this substitution must be removed.
            return false;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            // 0: USHORT LigatureGlyph
            writer.WriteUInt16BigEndian(this.LigatureGlyph);

            // 2: USHORT ComponentCount (must be Components.Length + 1, as the first glyph is implicit)
            writer.WriteUInt16BigEndian((ushort)(this.Components.Length + 1));

            // 4: USHORT[] ComponentGlyphIDs
            foreach (ushort componentGid in this.Components)
            {
                writer.WriteUInt16BigEndian(componentGid);
            }
        }
    }
}
