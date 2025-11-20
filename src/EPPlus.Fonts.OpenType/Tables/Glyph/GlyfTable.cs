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
using EPPlus.Fonts.OpenType.Tables.Loca;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    /// <summary>
    /// This table contains information that describes the glyphs in the font in the TrueType outline format
    /// https://docs.microsoft.com/en-us/typography/opentype/spec/glyf
    /// </summary>
    public class GlyfTable : FontTableBase
    {
        internal GlyfTable(TableLoaderSettings settings)
        {
            _tableLoaderSettings = settings;
        }

        internal GlyfTable(List<Glyph> glyphs)
        {
            Glyphs = glyphs ?? new List<Glyph>();
        }

        private readonly TableLoaderSettings _tableLoaderSettings;

        /// <summary>
        /// All glyphs in the font, indexed by glyph ID.
        /// </summary>
        public List<Glyph> Glyphs { get; set; }

        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            if (Glyphs == null || Glyphs.Count == 0)
                return;

            if (_tableLoaderSettings != null)
            {
                // Originalfont: use locaOffsets from loader
                var locaOffsets = _tableLoaderSettings.TableLoaders
                    .GetLocaTableLoader(_tableLoaderSettings)
                    .Load().Offsets;

                for (int i = 0; i < Glyphs.Count; i++)
                {
                    long glyphStart = writer.BaseStream.Position;
                    Glyph glyph = Glyphs[i];

                    if (glyph != null)
                        glyph.Serialize(writer);

                    long glyphEnd = writer.BaseStream.Position;
                    int writtenLength = (int)(glyphEnd - glyphStart);

                    if (i + 1 < locaOffsets.Count)
                    {
                        int expectedLength = (int)(locaOffsets[i + 1] - locaOffsets[i]);
                        if (expectedLength > writtenLength)
                        {
                            int padding = expectedLength - writtenLength;
                            for (int p = 0; p < padding; p++)
                                writer.Write((byte)0);
                        }
                    }
                }
            }
            else
            {
                // Subset-font: glyph handles its own padding
                foreach (var glyph in Glyphs)
                {
                    glyph.Serialize(writer); // Glyph.Serialize ska lägga till 4-byte alignment
                }
            }
        }

        internal override void Clear()
        {
            Glyphs.Clear();
        }

        public Glyph GetGlyph(ushort glyphId)
        {
            if (Glyphs == null || glyphId >= Glyphs.Count)
                return null;

            return Glyphs[glyphId];
        }


        public void ResolveCompositeGlyphs(HashSet<ushort> glyphSet)
        {
            bool addedNew;
            do
            {
                addedNew = false;
                foreach (var glyphId in glyphSet.ToList())
                {
                    var glyph = Glyphs[glyphId];
                    if (glyph.Header.numberOfContours < 0 && glyph.CompositeData != null)
                    {
                        foreach (var component in glyph.CompositeData.Components)
                        {
                            if (!glyphSet.Contains(component.GlyphIndex))
                            {
                                glyphSet.Add(component.GlyphIndex);
                                addedNew = true;
                            }
                        }
                    }
                }
            } while (addedNew);
        }

        public GlyfTable CreateSubset(List<ushort> sortedGlyphIds, Dictionary<ushort, ushort> idMapping)
        {
            var newGlyphs = new List<Glyph>(sortedGlyphIds.Count);

            foreach (var oldId in sortedGlyphIds)
            {
                var originalGlyph = Glyphs[oldId];
                var clonedGlyph = originalGlyph.Clone();

                // Remap composite glyph references
                if (clonedGlyph.Header.numberOfContours < 0 && clonedGlyph.CompositeData != null)
                {
                    foreach (var component in clonedGlyph.CompositeData.Components)
                    {
                        if (idMapping.TryGetValue(component.GlyphIndex, out ushort newIndex))
                        {
                            component.GlyphIndex = newIndex;
                        }
                        else
                        {
                            component.GlyphIndex = 0; // fallback to .notdef
                        }
                    }
                }

                newGlyphs.Add(clonedGlyph);
            }

            return new GlyfTable(newGlyphs);
        }

        public List<uint> CalculateOffsets()
        {
            var offsets = new List<uint>();
            uint currentOffset = 0;

            foreach (var glyph in Glyphs) // Glyphs är subset-listan i GlyfTable
            {
                offsets.Add(currentOffset);
                currentOffset += (uint)glyph.GetSize(); // GetLength() = antal bytes för glyfen
            }

            // Lägg till sista offset (slutet av sista glyfen)
            offsets.Add(currentOffset);

            return offsets;
        }
    }
}
