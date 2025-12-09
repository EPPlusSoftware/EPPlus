/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
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


        internal override int MapCodePointToGlyph(int codePoint)
        {
            foreach (var selector in VariationSelectors)
            {
                if (selector.NonDefaultUvsTable != null)
                {
                    foreach (var entry in selector.NonDefaultUvsTable.Mappings)
                    {
                        if (entry.UnicodeValue == codePoint)
                        {
                            return entry.GlyphId;
                        }
                    }
                }
            }
            return -1; // Not found
        }


        public override bool TryGetGlyphId(uint codePoint, out ushort glyphId)
        {
            glyphId = 0;

            foreach (var selector in VariationSelectors)
            {
                if (selector.NonDefaultUvsTable != null)
                {
                    foreach (var entry in selector.NonDefaultUvsTable.Mappings)
                    {
                        if (entry.UnicodeValue == codePoint)
                        {
                            glyphId = entry.GlyphId;
                            return glyphId != 0;
                        }
                    }
                }
            }

            return false;
        }



        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable14Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
