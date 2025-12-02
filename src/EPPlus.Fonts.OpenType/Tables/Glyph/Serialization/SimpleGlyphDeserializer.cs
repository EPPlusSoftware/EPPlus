using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Glyph.Serialization
{
    internal static class SimpleGlyphDeserializer
    {
        public static SimpleGlyph Deserialize(FontsBinaryReader reader, int numberOfContours)
        {
            var glyph = new SimpleGlyph();

            // 1. Read endPtsOfContours
            glyph.EndPtsOfContours = new ushort[numberOfContours];
            for (int i = 0; i < numberOfContours; i++)
            {
                glyph.EndPtsOfContours[i] = reader.ReadUInt16BigEndian();
            }

            // 2. Read instructions
            var instructionLength = reader.ReadUInt16BigEndian();
            glyph.Instructions = reader.ReadBytes(instructionLength);

            // 3. Calculate number of points
            int pointCount = glyph.EndPtsOfContours[numberOfContours - 1] + 1;


            // 4. Read flags (preserve original runs)
            var flagRuns = new List<FlagRun>();
            var expandedFlags = new List<byte>();

            while (expandedFlags.Count < pointCount)
            {
                byte flag = reader.ReadByte();
                byte repeatCount = 0;

                if ((flag & 0x08) != 0) // repeat flag
                {
                    repeatCount = reader.ReadByte();
                }

                flagRuns.Add(new FlagRun { Flag = flag, RepeatCount = repeatCount });

                // Expand for internal logic (coordinates)
                expandedFlags.Add(flag);
                for (int r = 0; r < repeatCount; r++)
                    expandedFlags.Add(flag);
            }

            glyph.FlagRuns = flagRuns;
            glyph.Flags = expandedFlags;


            // 5. Read X-coordinates as raw bytes
            var xBytes = new List<byte>();
            for (int i = 0; i < pointCount; i++)
            {
                if ((expandedFlags[i] & 0x02) != 0) // x-short
                {
                    xBytes.Add(reader.ReadByte());
                }
                else if ((expandedFlags[i] & 0x10) == 0) // not same
                {
                    xBytes.Add(reader.ReadByte());
                    xBytes.Add(reader.ReadByte());
                }
                // If same, no bytes written
            }
            glyph.XBytes = xBytes.ToArray();

            // 6. Read Y-coordinates as raw bytes
            var yBytes = new List<byte>();
            for (int i = 0; i < pointCount; i++)
            {
                if ((expandedFlags[i] & 0x04) != 0) // y-short
                {
                    yBytes.Add(reader.ReadByte());
                }
                else if ((expandedFlags[i] & 0x20) == 0) // not same
                {
                    yBytes.Add(reader.ReadByte());
                    yBytes.Add(reader.ReadByte());
                }
                // If same, no bytes written
            }
            glyph.YBytes = yBytes.ToArray();

            return glyph;
        }
    }
}
