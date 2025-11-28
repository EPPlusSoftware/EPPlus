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
                var locaOffsets = TableLoaders.GetLocaTableLoader(_tableLoaderSettings).Load().Offsets;
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
                // Subset-font: align every glyph to 4 bytes
                foreach (var glyph in Glyphs)
                {
                    long start = writer.BaseStream.Position;
                    glyph.Serialize(writer);
                    long end = writer.BaseStream.Position;

                    int writtenLength = (int)(end - start);
                    int padding = (4 - (writtenLength % 4)) % 4;
                    for (int p = 0; p < padding; p++)
                        writer.Write((byte)0);
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

    }
}
