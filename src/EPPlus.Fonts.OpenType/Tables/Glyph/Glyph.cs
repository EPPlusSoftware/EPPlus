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
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tables.Glyph
{
    public class Glyph : FontTableElement
    {
        public GlyphHeader Header { get; internal set; }

        public SimpleGlyph SimpleData { get; internal set; }

        public CompositeGlyph CompositeData { get; internal set; }


        public int GetSize()
        {
            using (var ms = new MemoryStream())
            using (var writer = new FontsBinaryWriter(ms))
            {
                Serialize(writer);
                return (int)ms.Length;
            }
        }



        internal override void Serialize(FontsBinaryWriter writer)
        {
            // Spara startpositionen
            long start = writer.BaseStream.Position;

            // Skriv header
            Header.Serialize(writer);

            // Skriv glyfdata beroende på typ
            if (Header.numberOfContours > 0 && SimpleData != null)
            {
                SimpleData.Serialize(writer);
            }
            else if (Header.numberOfContours < 0 && CompositeData != null)
            {
                CompositeData.Serialize(writer);
            }
            // Om numberOfContours == 0 → tom glyf, bara header

            // Lägg till padding till 4-byte boundary
            long end = writer.BaseStream.Position;
            int writtenLength = (int)(end - start);
            int padding = (4 - (writtenLength % 4)) % 4;
            for (int p = 0; p < padding; p++)
                writer.Write((byte)0);
        }


        public Glyph Clone()
        {
            var clone = new Glyph
            {
                Header = new GlyphHeader
                {
                    numberOfContours = this.Header.numberOfContours,
                    xMin = this.Header.xMin,
                    yMin = this.Header.yMin,
                    xMax = this.Header.xMax,
                    yMax = this.Header.yMax
                }
            };

            // Clone SimpleGlyph if present
            if (this.SimpleData != null)
            {
                clone.SimpleData = new SimpleGlyph
                {
                    EndPtsOfContours = (ushort[])this.SimpleData.EndPtsOfContours.Clone(),
                    Instructions = (byte[])this.SimpleData.Instructions.Clone(),
                    XBytes = (byte[])this.SimpleData.XBytes.Clone(),
                    YBytes = (byte[])this.SimpleData.YBytes.Clone(),
                    Flags = new List<byte>(this.SimpleData.Flags),
                    FlagRuns = this.SimpleData.FlagRuns
                        .Select(fr => new FlagRun { Flag = fr.Flag, RepeatCount = fr.RepeatCount })
                        .ToList(),
                    Points = this.SimpleData.Points
                        .Select(p => new GlyphPoint { X = p.X, Y = p.Y, OnCurve = p.OnCurve })
                        .ToList()
                };
            }

            // Clone CompositeGlyph if present
            if (this.CompositeData != null)
            {
                clone.CompositeData = new CompositeGlyph
                {
                    Instructions = (byte[])this.CompositeData.Instructions.Clone(),
                    Components = this.CompositeData.Components
                        .Select(c => new GlyphComponent
                        {
                            Flags = c.Flags,
                            GlyphIndex = c.GlyphIndex, // Will be remapped later in GlyfTable.CreateSubset
                            Argument1 = c.Argument1,
                            Argument2 = c.Argument2,
                            Scale = c.Scale,
                            XScale = c.XScale,
                            YScale = c.YScale,
                            Scale01 = c.Scale01,
                            Scale10 = c.Scale10
                        })
                        .ToList()
                };
            }

            return clone;
        }

    }
}
