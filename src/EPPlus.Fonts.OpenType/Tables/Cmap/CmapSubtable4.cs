using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Cmap.Serialization;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Cmap
{
    public class CmapSubtable4 : CmapSubtableBase
    {

        public override ushort Format { get; } = 4;

        public override uint Length { get; internal set; }

        public override  uint Language { get; internal set; }

        public ushort SegCountX2 { get; internal set; }
        public ushort SearchRange { get; internal set; }
        public ushort EntrySelector { get; internal set; }
        public ushort RangeShift { get; internal set; }

        public ushort[] EndCode { get; internal set; } = new ushort[0];
        public ushort ReservedPad { get; internal set; }
        public ushort[] StartCode { get; internal set; } = new ushort[0];
        public short[] IdDelta { get; internal set; } = new short[0];
        public ushort[] IdRangeOffset { get; internal set; } = new ushort[0];

        public ushort[] GlyphIdArray { get; internal set; } = new ushort[0];

        public override GlyphMappings GetGlyphMappings()
        {
            var mapping = new GlyphMappings();

            int segCount = EndCode.Length;

            for (int i = 0; i < segCount; i++)
            {
                ushort startCode = StartCode[i];
                ushort endCode = EndCode[i];
                short idDelta = IdDelta[i];
                ushort idRangeOffset = IdRangeOffset[i];

                for (uint charCode = startCode; charCode <= endCode; charCode++)
                {
                    ushort glyphIndex;

                    if (idRangeOffset == 0)
                    {
                        glyphIndex = (ushort)((charCode + idDelta) % 65536);
                    }
                    else
                    {
                        int offsetIndex = (idRangeOffset / 2) + (int)(charCode - startCode) - (segCount - i);
                        if (offsetIndex >= 0 && offsetIndex < GlyphIdArray.Length)
                        {
                            ushort glyphId = GlyphIdArray[offsetIndex];
                            if (glyphId != 0)
                            {
                                glyphIndex = (ushort)((glyphId + idDelta) % 65536);
                            }
                            else
                            {
                                glyphIndex = 0;
                            }
                        }
                        else
                        {
                            glyphIndex = 0;
                        }
                    }

                    if (glyphIndex != 0)
                    {
                        mapping.AddMapping(charCode, glyphIndex);
                    }
                }
            }

            return mapping;
        }


        internal override void Serialize(FontsBinaryWriter writer)
        {
            var serializer = new CmapSubtable4Serializer();
            serializer.Serialize(this, writer);
        }
    }
}
