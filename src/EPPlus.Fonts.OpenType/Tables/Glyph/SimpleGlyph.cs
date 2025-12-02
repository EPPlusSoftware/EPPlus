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

using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    public class SimpleGlyph : FontTableElement
    {
        public ushort[] EndPtsOfContours { get; set; }
        public byte[] Instructions { get; set; }
        public List<GlyphPoint> Points { get; set; } = new();

        public List<byte> Flags { get; set; } = new List<byte>();

        public List<FlagRun> FlagRuns { get; set; } = new List<FlagRun>();


        // Raw encoded coordinate bytes for X and Y
        public byte[] XBytes { get; set; }
        public byte[] YBytes { get; set; }


        /// <summary>
        /// Serialize according to OpenType "Simple glyph description".
        /// </summary>
        internal override void Serialize(FontsBinaryWriter writer)
        {


            // Write endPtsOfContours
            foreach (var endPt in EndPtsOfContours)
                writer.WriteUInt16BigEndian(endPt);

            // Write instructions
            writer.WriteUInt16BigEndian((ushort)Instructions.Length);
            writer.Write(Instructions);

            // Write flags using original runs
            foreach (var run in FlagRuns)
            {
                writer.Write(run.Flag);
                if ((run.Flag & 0x08) != 0)
                {
                    writer.Write(run.RepeatCount);
                }
            }

            // Write coordinates
            writer.Write(XBytes);
            writer.Write(YBytes);


        }
    }
}
