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

            // 4. Read flags (needed to know how to parse coordinates)
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

                expandedFlags.Add(flag);
                for (int r = 0; r < repeatCount; r++)
                    expandedFlags.Add(flag);
            }

            glyph.FlagRuns = flagRuns;
            glyph.Flags = expandedFlags;

            // 5. Read X-coordinates by measuring the byte block
            // We must do this because TrueType uses a delta-encoding where 
            // some points consume 0, 1, or 2 bytes depending on flags.
            long xStart = reader.BaseStream.Position;
            for (int i = 0; i < pointCount; i++)
            {
                byte flag = expandedFlags[i];
                if ((flag & 0x02) != 0) // X-Short
                {
                    reader.ReadByte();
                }
                else if ((flag & 0x10) == 0) // Not same (consumes 2 bytes)
                {
                    reader.ReadInt16BigEndian();
                }
                // If (flag & 0x10) is set and X-Short is NOT set, 0 bytes are consumed (Same as prev)
            }
            long xEnd = reader.BaseStream.Position;
            int xLength = (int)(xEnd - xStart);

            // Go back and grab the raw bytes
            reader.BaseStream.Position = xStart;
            glyph.XBytes = reader.ReadBytes(xLength);

            // 6. Read Y-coordinates by measuring the byte block
            long yStart = reader.BaseStream.Position;
            for (int i = 0; i < pointCount; i++)
            {
                byte flag = expandedFlags[i];
                if ((flag & 0x04) != 0) // Y-Short
                {
                    reader.ReadByte();
                }
                else if ((flag & 0x20) == 0) // Not same (consumes 2 bytes)
                {
                    reader.ReadInt16BigEndian();
                }
            }
            long yEnd = reader.BaseStream.Position;
            int yLength = (int)(yEnd - yStart);

            reader.BaseStream.Position = yStart;
            glyph.YBytes = reader.ReadBytes(yLength);

            return glyph;
        }
    }
}
