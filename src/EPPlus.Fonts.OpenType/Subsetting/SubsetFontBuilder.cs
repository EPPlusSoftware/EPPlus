using System;
using System.Collections.Generic;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Loca;
using EPPlus.Fonts.OpenType.Tables.Maxp;
using EPPlus.Fonts.OpenType.Tables.Os2;
using EPPlus.Fonts.OpenType.Tables.Post;

namespace EPPlus.Fonts.OpenType.Subsetting
{
    internal class SubsetFontBuilder
    {
        public OpenTypeFont CreateSubset(OpenTypeFont originalFont, IEnumerable<int> unicodeChars)
        {
            // Always include space (U+0020) – required for OS/2 usDefaultChar/usBreakChar
            int[] inputChars = unicodeChars as int[] ?? new List<int>(unicodeChars).ToArray();
            List<int> extendedChars = new List<int>(inputChars);
            if (!extendedChars.Contains(32))
                extendedChars.Add(32);

            // 1. Build initial glyph set
            HashSet<ushort> glyphSet = BuildGlyphSet(originalFont, extendedChars);

            // 2. Resolve composite glyphs ON THE ORIGINAL FONT – critical!
            //originalFont.GlyfTable.ResolveCompositeGlyphs(glyphSet);
            glyphSet.RemoveWhere(gid =>
            {
                string name = originalFont.GlyfTable.GetGlyphName(gid, originalFont);
                return name == "uni03BC" || name == "mu" || name == "lambda" || name == "sigma1";
            });

            // 3. Create subset font
            OpenTypeFont newFont = new OpenTypeFont(originalFont.Format);

            // head
            newFont.AddOrReplaceTable(originalFont.HeadTable.Clone());

            // name
            if (originalFont.NameTable != null)
                newFont.AddOrReplaceTable(originalFont.NameTable.Clone());

            // maxp & hhea – clone now, update later
            MaxpTable maxpTable = originalFont.MaxpTable != null ? originalFont.MaxpTable.Clone() : null;
            HheaTable hheaTable = originalFont.HheaTable != null ? originalFont.HheaTable.Clone() : null;

            // hmtx
            if (originalFont.HmtxTable != null)
            {
                var hmtx = originalFont.HmtxTable.CloneSubset(glyphSet, originalFont.HmtxTable);
                newFont.AddOrReplaceTable(hmtx);
            }

            // 4. glyf + loca
            Dictionary<ushort, ushort> oldToNewGlyphId = new Dictionary<ushort, ushort>();
            List<ushort> newToOldGlyphId = new List<ushort>();

            SubsetGlyfAndLoca(newFont, originalFont, glyphSet, oldToNewGlyphId, newToOldGlyphId);

            // 5. Update maxp and hhea
            int finalGlyphCount = newToOldGlyphId.Count;

            if (maxpTable != null)
            {
                maxpTable.numGlyphs = (ushort)finalGlyphCount;
                newFont.AddOrReplaceTable(maxpTable);
            }

            if (hheaTable != null)
            {
                hheaTable.numberOfHMetrics = (ushort)finalGlyphCount;
                newFont.AddOrReplaceTable(hheaTable);
            }

            if (originalFont.HmtxTable != null)
            {
                var finalHmtx = originalFont.HmtxTable.CloneForGlyphCount(
                    newToOldGlyphId.Count,
                    originalFont.MaxpTable.numGlyphs);

                int originalHMetricsCount = originalFont.HmtxTable.hMetrics.Count;

                for (int i = 0; i < newToOldGlyphId.Count; i++)
                {
                    ushort oldId = newToOldGlyphId[i];

                    // Hämta från original
                    ushort advanceWidth = originalFont.HmtxTable.GetAdvanceWidth(oldId);
                    short lsb = originalFont.HmtxTable.GetLeftSideBearing(oldId);

                    // Sätt alltid advanceWidth i hMetrics
                    finalHmtx.hMetrics[i].advanceWidth = advanceWidth;

                    // Sätt LSB – men bara på rätt plats
                    if (i < originalHMetricsCount)
                    {
                        // Inom hMetrics-området – LSB finns i hMetrics[i].lsb
                        finalHmtx.hMetrics[i].lsb = lsb;
                    }
                    else
                    {
                        // Utanför – LSB finns i leftSideBearings
                        int lsbIndex = i - originalHMetricsCount;
                        if (lsbIndex < finalHmtx.leftSideBearings.Count)
                        {
                            finalHmtx.leftSideBearings[lsbIndex] = lsb;
                        }
                        // Om listan är för kort – lägg till 0 (säkert)
                        else if (lsbIndex == finalHmtx.leftSideBearings.Count)
                        {
                            finalHmtx.leftSideBearings.Add(lsb);
                        }
                    }
                }

                newFont.AddOrReplaceTable(finalHmtx);
            }

            // 6. cmap
            SubsetCmap(newFont, originalFont, oldToNewGlyphId);

            // 7. OS/2
            SubsetOS2(newFont, originalFont);

            // 8. post
            SubsetPost(newFont, originalFont);

            return newFont;
        }

        private HashSet<ushort> BuildGlyphSet(OpenTypeFont font, IEnumerable<int> unicodeChars)
        {
            HashSet<ushort> glyphIds = new HashSet<ushort>();
            foreach (int codePoint in unicodeChars)
            {
                ushort glyphId;
                if (font.CmapTable.TryGetGlyphId((uint)codePoint, out glyphId))
                {
                    glyphIds.Add(glyphId);
                }
            }
            return glyphIds;
        }

        private void SubsetGlyfAndLoca(
            OpenTypeFont newFont,
            OpenTypeFont originalFont,
            HashSet<ushort> glyphSet,
            Dictionary<ushort, ushort> oldToNewGlyphId,
            List<ushort> newToOldGlyphId)
        {
            GlyfTable originalGlyf = originalFont.GlyfTable;

            // Ensure .notdef is first
            List<ushort> sortedIds = new List<ushort>(glyphSet);
            sortedIds.Sort();
            if (!sortedIds.Contains(0))
                sortedIds.Insert(0, 0);

            // Build mapping
            oldToNewGlyphId.Clear();
            newToOldGlyphId.Clear();
            for (int i = 0; i < sortedIds.Count; i++)
            {
                ushort oldId = sortedIds[i];
                oldToNewGlyphId[oldId] = (ushort)i;
                newToOldGlyphId.Add(oldId);
            }

            // Clone glyphs
            List<Glyph> newGlyphs = new List<Glyph>();
            foreach (ushort oldId in sortedIds)
            {
                Glyph oldGlyph = originalGlyf.GetGlyph(oldId);
                if (oldGlyph == null || oldGlyph.Header.numberOfContours == 0)
                {
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

            newFont.AddOrReplaceTable(new GlyfTable(newGlyphs));

            // Build loca
            List<uint> offsets = new List<uint> { 0 };
            uint current = 0;
            foreach (Glyph g in newGlyphs)
            {
                int size = g.GetSize();
                current += (uint)((size + 3) & ~3);
                offsets.Add(current);
            }

            bool shortFormat = true;
            foreach (uint o in offsets)
            {
                if (o > 131070)
                {
                    shortFormat = false;
                    break;
                }
            }

            newFont.HeadTable.IndexToLocFormat = shortFormat
                ? HeadTable.IndexToLocFormats.Offset16
                : HeadTable.IndexToLocFormats.Offset32;

            newFont.AddOrReplaceTable(LocaTable.CreateSubset(offsets, newFont.HeadTable.IndexToLocFormat));
        }

        private Glyph CloneGlyphWithRemappedComponents(Glyph original, Dictionary<ushort, ushort> oldToNew)
        {
            short xMin = original.Header.xMin;
            short xMax = original.Header.xMax;
            short yMin = original.Header.yMin;
            short yMax = original.Header.yMax;

            // Fix inverted bounding boxes (common in .notdef and some broken fonts)
            if (xMin > xMax)
            {
                short temp = xMin;
                xMin = xMax;
                xMax = temp;
            }
            if (yMin > yMax)
            {
                short temp = yMin;
                yMin = yMax;
                yMax = temp;
            }

            GlyphHeader header = new GlyphHeader(
                original.Header.numberOfContours,
                new BoundingRectangle(xMin, yMin, xMax, yMax));

            // Simple glyph – just clone
            if (original.Header.numberOfContours > 0 && original.SimpleData != null)
            {
                SimpleGlyph simple = new SimpleGlyph
                {
                    EndPtsOfContours = (ushort[])original.SimpleData.EndPtsOfContours.Clone(),
                    Instructions = (byte[])original.SimpleData.Instructions.Clone(),
                    XBytes = (byte[])original.SimpleData.XBytes.Clone(),
                    YBytes = (byte[])original.SimpleData.YBytes.Clone(),
                    FlagRuns = new List<FlagRun>(original.SimpleData.FlagRuns)
                };
                return new Glyph { Header = header, SimpleData = simple };
            }

            // Composite glyph – remap component glyph IDs
            if (original.Header.numberOfContours < 0 && original.CompositeData != null)
            {
                CompositeGlyph composite = new CompositeGlyph
                {
                    Instructions = (byte[])original.CompositeData.Instructions.Clone(),
                    Components = new List<GlyphComponent>()
                };

                foreach (GlyphComponent comp in original.CompositeData.Components)
                {
                    ushort newGlyphIndex = 0; // fallback to .notdef
                    oldToNew.TryGetValue(comp.GlyphIndex, out newGlyphIndex);

                    GlyphComponent newComp = new GlyphComponent
                    {
                        Flags = comp.Flags,
                        GlyphIndex = newGlyphIndex,
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

            // Empty glyph (like space)
            return new Glyph { Header = header };
        }

        private void SubsetCmap(OpenTypeFont newFont, OpenTypeFont originalFont, Dictionary<ushort, ushort> oldToNewGlyphId)
        {
            CmapTable originalCmap = originalFont.CmapTable;
            if (originalCmap == null || originalCmap.SubTables == null || originalCmap.SubTables.Count == 0)
                return;

            Dictionary<uint, ushort> needed = new Dictionary<uint, ushort>();

            foreach (CmapSubtableBase sub in originalCmap.SubTables)
            {
                GlyphMappings map = sub.GetGlyphMappings();
                if (map == null || map.CharCodeToGlyphIndex == null) continue;

                foreach (KeyValuePair<uint, ushort> kvp in map.CharCodeToGlyphIndex)
                {
                    ushort newId;
                    if (oldToNewGlyphId.TryGetValue(kvp.Value, out newId))
                    {
                        needed[kvp.Key] = newId;
                    }
                }
            }
            needed[0] = 0; // .notdef

            CmapSubtable4 format4 = CmapFormat4.CreateFromMappings(needed);

            CmapTable newCmap = new CmapTable();

            newCmap.EncodingRecords.Add(new EncodingRecord(Platforms.Windows, 1, 0));
            newCmap.EncodingRecords.Add(new EncodingRecord(Platforms.Unicode, 3, 0));

            newCmap.EncodingRecords[0].Subtable = format4;
            newCmap.EncodingRecords[1].Subtable = format4;
            newCmap.SubTables.Add(format4);

            newFont.AddOrReplaceTable(newCmap);
        }

        private void SubsetOS2(OpenTypeFont newFont, OpenTypeFont originalFont)
        {
            Os2Table original = originalFont.Os2Table;
            if (original == null) return;

            Os2Table os2 = original.Clone();

            os2.usDefaultChar = 32;
            os2.usBreakChar = 32;

            short maxAscent = short.MinValue;
            short maxDescent = short.MaxValue;

            foreach (Glyph g in newFont.GlyfTable.Glyphs)
            {
                if (g != null && g.Header != null)
                {
                    if (g.Header.yMax > maxAscent) maxAscent = g.Header.yMax;
                    if (g.Header.yMin < maxDescent) maxDescent = g.Header.yMin;
                }
            }

            if (maxAscent > os2.usWinAscent) os2.usWinAscent = (ushort)maxAscent;
            if (-maxDescent > os2.usWinDescent) os2.usWinDescent = (ushort)(-maxDescent);

            HheaTable hhea = newFont.HheaTable;
            if (hhea != null)
            {
                hhea.ascender = os2.sTypoAscender;
                hhea.descender = os2.sTypoDescender;
                hhea.lineGap = os2.sTypoLineGap;
            }

            newFont.AddOrReplaceTable(os2);
        }

        private void SubsetPost(OpenTypeFont newFont, OpenTypeFont originalFont)
        {
            if (originalFont.PostTable == null) return;

            // Skapa en ny, säker post-tabell med format 3.0
            // Format 3.0 = "No glyph names" – helt tillåtet och rekommenderat för subset-fonts
            var post = new PostTable
            {
                version = new Version16Dot16(0x00030000),  // 3.0 i packed format
                italicAngle = originalFont.PostTable.italicAngle,
                underlinePosition = originalFont.PostTable.underlinePosition,
                underlineThickness = originalFont.PostTable.underlineThickness,
                isFixedPitch = originalFont.PostTable.isFixedPitch,
                minMemType42 = 0,
                maxMemType42 = 0,
                minMemType1 = 0,
                maxMemType1 = 0
            };

            // Rensa bort gamla namn-data (om de finns)
            // De här fälten finns bara i format 2.0 – vi sätter dem till null för säkerhets skull
            post.glyphNameIndex = null;
            post.glyphNames = null;

            newFont.AddOrReplaceTable(post);
        }
    }
}