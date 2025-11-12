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

        private readonly TableLoaderSettings _tableLoaderSettings;

        /// <summary>
        /// All glyphs in the font, indexed by glyph ID.
        /// </summary>
        public Glyph[] Glyphs { get; set; }

        internal override void SerializeInternal(FontsBinaryWriter writer)
        {
            if (Glyphs == null || Glyphs.Length == 0)
                return;
            var locaOffsets = TableLoaders.GetLocaTableLoader(_tableLoaderSettings).Load().Offsets;
            long startPosition = writer.BaseStream.Position;

            for (int i = 0; i < Glyphs.Length; i++)
            {
                long glyphStart = writer.BaseStream.Position;
                Glyph glyph = Glyphs[i];

                if (glyph != null)
                    glyph.Serialize(writer);

                long glyphEnd = writer.BaseStream.Position;
                int writtenLength = (int)(glyphEnd - glyphStart);

                if (i + 1 < locaOffsets.Length)
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
    }
}
