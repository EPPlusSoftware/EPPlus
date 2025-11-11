using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    internal class CmapSubtable14 : CmapSubtableBase
    {
        public override ushort Format { get; } = 14;

        public override uint Length { get; internal set; }

        public override uint Language { get; internal set; }

        public List<VariationSelector> VariationSelectors { get; } = new List<VariationSelector>();

        public override GlyphMappings GetGlyphMappings()
        {
            var mapping = new GlyphMappings();

            // Iterate through all variation selectors in the table
            foreach (var selector in VariationSelectors)
            {
                // Only Non-Default UVS tables contain explicit glyph mappings
                if (selector.NonDefaultUvsTable != null)
                {
                    foreach (var entry in selector.NonDefaultUvsTable.Mappings)
                    {
                        // Add mapping from Unicode value to glyph index
                        // Note: Variation selector is ignored here, since GlyphMapping is flat
                        mapping.AddMapping(entry.UnicodeValue, entry.GlyphId);
                    }
                }

                // Default UVS tables do not contain glyph indices and are not included
            }

            return mapping;
        }

        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable14Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
