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
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Loca;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class SubsetFontBuilder
    {

        public OpenTypeFont CreateSubset(OpenTypeFont originalFont, IEnumerable<int> unicodeChars)
        {
            var glyphSet = BuildGlyphSet(originalFont, unicodeChars); // HashSet<ushort>

            // 1) head
            var newFont = new OpenTypeFont(originalFont.Format);
            var headTable = originalFont.HeadTable.Clone();
            newFont.AddOrReplaceTable(headTable);


            // 2) name
            if (originalFont.NameTable != null)
            {
                var nameTable = originalFont.NameTable.Clone();
                newFont.AddOrReplaceTable(nameTable);
            }

            // 3) maxp
            if (originalFont.MaxpTable != null)
            {
                var maxpTable = originalFont.MaxpTable.Clone();
                maxpTable.numGlyphs = (ushort)glyphSet.Count; // glyphSet from BuildGlyphSet
                newFont.AddOrReplaceTable(maxpTable);
            }

            // 4) hhea

            if (originalFont.HheaTable != null)
            {
                var hheaTable = originalFont.HheaTable.Clone();
                hheaTable.numberOfHMetrics = (ushort)glyphSet.Count; // Temporärt, tills hmtx är klar
                newFont.AddOrReplaceTable(hheaTable);
            }

            // 5) hmtx
            if (originalFont.HmtxTable != null)
            {
                var hmtxTable = originalFont.HmtxTable.CloneSubset(glyphSet, originalFont.HmtxTable);
                newFont.AddOrReplaceTable(hmtxTable);
            }

            // 6 glyf an loca
            var oldToNewGlyphId = new Dictionary<ushort, ushort>();
            var newToOldGlyphId = new List<ushort>(); // index = ny glyph ID

            SubsetGlyfAndLoca(newFont, originalFont, glyphSet, oldToNewGlyphId, newToOldGlyphId);


            return newFont;

        }


        private HashSet<ushort> BuildGlyphSet(OpenTypeFont font, IEnumerable<int> unicodeChars)
        {
            var glyphIds = new HashSet<ushort>();

            foreach (var codePoint in unicodeChars)
            {
                if (font.CmapTable.TryGetGlyphId(codePoint, out ushort glyphId))
                {
                    glyphIds.Add(glyphId);
                }
            }

            glyphIds.Add(0); // Always include .notdef

            font.GlyfTable.ResolveCompositeGlyphs(glyphIds);

            return glyphIds;
        }

        private void SubsetGlyfAndLoca(
            OpenTypeFont newFont,
            OpenTypeFont originalFont,
            HashSet<ushort> glyphSet,
            Dictionary<ushort, ushort> oldToNewGlyphId,
            List<ushort> newToOldGlyphId)
        {
            var originalGlyf = originalFont.GlyfTable;
            var originalLoca = originalFont.LocaTable;
            var originalHead = originalFont.HeadTable;
            var originalMaxp = originalFont.MaxpTable;

            if (originalGlyf == null || originalLoca == null || originalHead == null || originalMaxp == null)
                throw new InvalidOperationException("Font missing required tables for subsetting (glyf, loca, head, maxp).");

            // 1. Bygg sorterad lista av glyphs vi vill ha med
            var sortedOldGlyphIds = glyphSet.OrderBy(g => g).ToList();

            // Säkerställ att .notdef (0) är först – enligt TrueType-spec
            if (!sortedOldGlyphIds.Contains(0))
                sortedOldGlyphIds.Insert(0, 0);

            // 2. Bygg remapping: old ID → new ID
            oldToNewGlyphId.Clear();
            newToOldGlyphId.Clear();

            for (int newId = 0; newId < sortedOldGlyphIds.Count; newId++)
            {
                ushort oldId = sortedOldGlyphIds[newId];
                oldToNewGlyphId[oldId] = (ushort)newId;
                newToOldGlyphId.Add(oldId);
            }

            // 3. Klona glyphs med remappade composite-referenser
            var newGlyphs = new List<Glyph>(sortedOldGlyphIds.Count);

            foreach (ushort oldId in sortedOldGlyphIds)
            {
                var oldGlyph = originalGlyf.GetGlyph(oldId);

                if (oldGlyph == null || oldGlyph.Header.numberOfContours == 0)
                {
                    // Tom glyph (t.ex. space, .null) → tom header
                    newGlyphs.Add(new Glyph
                    {
                        Header = new GlyphHeader(0, new BoundingRectangle(0, 0, 0, 0))
                    });
                }
                else
                {
                    newGlyphs.Add(CloneGlyphWithRemappedComponents(oldGlyph, oldToNewGlyphId));
                }
            }

            // 4. Skapa ny GlyfTable
            var newGlyfTable = new GlyfTable(newGlyphs);
            newFont.AddOrReplaceTable(newGlyfTable);

            // 5. Beräkna nya offsets och välj bästa loca-format
            var locaOffsets = new List<uint> { 0 };
            uint currentOffset = 0;

            foreach (var glyph in newGlyphs)
            {
                int rawSize = glyph.GetSize();
                int paddedSize = (rawSize + 3) & ~3; // 4-byte align
                currentOffset += (uint)paddedSize;
                locaOffsets.Add(currentOffset);
            }

            // Välj format: short om alla offsets ≤ 131070 (eftersom /2)
            bool useShortFormat = locaOffsets.All(o => o <= 131070);
            var indexToLocFormat = useShortFormat
                ? HeadTable.IndexToLocFormats.Offset16
                : HeadTable.IndexToLocFormats.Offset32;

            // Uppdatera head.indexToLocFormat
            newFont.HeadTable.IndexToLocFormat = indexToLocFormat;

            // 6. Skapa ny LocaTable med din exakta factory
            var newLocaTable = LocaTable.CreateSubset(locaOffsets, indexToLocFormat, newFont.MaxpTable);
            newFont.AddOrReplaceTable(newLocaTable);
        }

        private Glyph CloneGlyphWithRemappedComponents(Glyph original, Dictionary<ushort, ushort> oldToNew)
        {
            var header = new GlyphHeader(original.Header.numberOfContours,
                new BoundingRectangle(original.Header.xMin, original.Header.xMax,
                                      original.Header.yMin, original.Header.yMax));

            if (original.Header.numberOfContours > 0 && original.SimpleData != null)
            {
                // Simple glyph – klona rakt av
                var simple = new SimpleGlyph
                {
                    EndPtsOfContours = (ushort[])original.SimpleData.EndPtsOfContours.Clone(),
                    Instructions = (byte[])original.SimpleData.Instructions.Clone(),
                    XBytes = (byte[])original.SimpleData.XBytes.Clone(),
                    YBytes = (byte[])original.SimpleData.YBytes.Clone(),
                    FlagRuns = new List<FlagRun>(original.SimpleData.FlagRuns)
                };
                return new Glyph { Header = header, SimpleData = simple };
            }
            else if (original.Header.numberOfContours < 0 && original.CompositeData != null)
            {
                // Composite – remappa alla komponenters GlyphIndex!
                var composite = new CompositeGlyph
                {
                    Instructions = (byte[])original.CompositeData.Instructions.Clone(),
                    Components = new List<GlyphComponent>()
                };

                foreach (var comp in original.CompositeData.Components)
                {
                    var newComp = new GlyphComponent
                    {
                        Flags = comp.Flags,
                        GlyphIndex = oldToNew.ContainsKey(comp.GlyphIndex)
                            ? oldToNew[comp.GlyphIndex]
                            : (ushort)0, // fallback till .notdef
                        Argument1 = comp.Argument1,
                        Argument2 = comp.Argument2,
                        Scale = comp.Scale,
                        XScale = comp.XScale,
                        YScale = comp.YScale,
                        Scale01 = comp.Scale01,
                        Scale10 = comp.Scale10
                    };
                    composite.Components.Add(newComp);
                }

                return new Glyph { Header = header, CompositeData = composite };
            }

            // Tom glyph
            return new Glyph { Header = header };
        }

        private uint[] BuildLocaOffsets(List<Glyph> glyphs, bool useLongFormat)
        {
            var offsets = new List<uint> { 0 };
            uint current = 0;

            foreach (var glyph in glyphs)
            {
                int size = glyph.GetSize();
                int padded = (size + 3) & ~3; // 4-byte align
                current += (uint)padded;
                offsets.Add(current);
            }

            if (!useLongFormat)
            {
                // Short format: dela med 2
                return offsets.Select(o => o / 2).ToArray();
            }

            return offsets.ToArray();
        }
    }
}
