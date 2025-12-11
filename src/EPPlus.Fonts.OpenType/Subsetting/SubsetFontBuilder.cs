using System;
using System.Collections.Generic;
using System.Linq;
using EPPlus.Fonts.OpenType.Tables;
using EPPlus.Fonts.OpenType.Tables.Cmap;
using EPPlus.Fonts.OpenType.Tables.Cmap.Mappings;
using EPPlus.Fonts.OpenType.Tables.Glyph;
using EPPlus.Fonts.OpenType.Tables.Head;
using EPPlus.Fonts.OpenType.Tables.Hhea;
using EPPlus.Fonts.OpenType.Tables.Kern;
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
            // Konvertera och spara alla använda code points (detta är nyckeln!)
            var usedCodePoints = unicodeChars
                .Select(c => (uint)c)
                .Distinct()
                .ToList();

            // Always include space
            if (!usedCodePoints.Contains(32))
                usedCodePoints.Add(32);

            // 1. Bygg initial glyph set från de faktiska tecknen
            var glyphSet = new HashSet<ushort>();
            foreach (uint cp in usedCodePoints)
            {
                if (originalFont.CmapTable.TryGetGlyphId(cp, out ushort gid))
                    glyphSet.Add(gid);
            }

            // 2. Resolve composite glyphs – nu lägger vi till alla komponenter
            originalFont.GlyfTable.ResolveCompositeGlyphs(glyphSet);

            // Säkerställ .notdef (GID 0) finns och ska bli ny GID 0
            glyphSet.Add(0);

            // 3. Skapa subset-font
            var newFont = new OpenTypeFont(originalFont.Format);

            // head + name
            newFont.AddOrReplaceTable(originalFont.HeadTable.Clone());
            if (originalFont.NameTable != null)
                newFont.AddOrReplaceTable(originalFont.NameTable.Clone());

            // maxp & hhea
            var maxpTable = originalFont.MaxpTable?.Clone();
            var hheaTable = originalFont.HheaTable?.Clone();

            // hmtx (preliminär)
            if (originalFont.HmtxTable != null)
                newFont.AddOrReplaceTable(originalFont.HmtxTable.CloneSubset(glyphSet, originalFont.HmtxTable));

            // 4. glyf + loca + mappning
            var oldToNewGlyphId = new Dictionary<ushort, ushort>();
            var newToOldGlyphId = new List<ushort>();

            SubsetGlyfAndLoca(newFont, originalFont, glyphSet, oldToNewGlyphId, newToOldGlyphId);

            // 5. Uppdatera maxp och hhea
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

            // 6. Final hmtx med korrekt antal glyfer
            if (originalFont.HmtxTable != null)
            {
                var finalHmtx = originalFont.HmtxTable.CloneForGlyphCount(finalGlyphCount, originalFont.MaxpTable.numGlyphs);
                int originalHMetricsCount = originalFont.HmtxTable.hMetrics.Count;

                for (int i = 0; i < finalGlyphCount; i++)
                {
                    ushort oldId = newToOldGlyphId[i];
                    ushort advance = originalFont.HmtxTable.GetAdvanceWidth(oldId);
                    short lsb = originalFont.HmtxTable.GetLeftSideBearing(oldId);

                    finalHmtx.hMetrics[i].advanceWidth = advance;
                    if (i < originalHMetricsCount)
                        finalHmtx.hMetrics[i].lsb = lsb;
                    else
                        finalHmtx.leftSideBearings[i - originalHMetricsCount] = lsb;
                }
                newFont.AddOrReplaceTable(finalHmtx);
            }

            // 7. Bygg cmap direkt från de använda tecknen – DETTA ÄR FIXEN!
            var cmapMapping = new Dictionary<uint, ushort>();
            foreach (uint cp in usedCodePoints)
            {
                if (originalFont.CmapTable.TryGetGlyphId(cp, out ushort oldGid) &&
                    oldToNewGlyphId.TryGetValue(oldGid, out ushort newGid))
                {
                    cmapMapping[cp] = newGid;
                }
                else
                {
                    cmapMapping[cp] = 0; // .notdef
                }
            }
            cmapMapping[0] = 0; // .notdef för kodpunkt 0

            var format4 = CmapFormat4.CreateFromMappings(cmapMapping);

            var newCmap = new CmapTable();
            newCmap.EncodingRecords.Add(new EncodingRecord(Platforms.Windows, 1, 0));
            newCmap.EncodingRecords.Add(new EncodingRecord(Platforms.Unicode, 3, 0));
            newCmap.SubTables.Add(format4);
            newCmap.EncodingRecords[0].Subtable = format4;
            newCmap.EncodingRecords[1].Subtable = format4;
            newFont.AddOrReplaceTable(newCmap);

            // 8. OS/2 och post
            var usedCodePointsSet = new HashSet<uint>(usedCodePoints);
            SubsetOS2(newFont, originalFont, usedCodePointsSet);
            SubsetPost(newFont, originalFont);

            // 9. Kern-tabellen – behåll alla par som finns i vår glyphSet
            SubsetKern(newFont, originalFont, oldToNewGlyphId);

            // Spara för eventuell debug (valfritt)
            newFont.UsedCodePointsForSubset = usedCodePoints.Select(c => c).ToList();

            return newFont;
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

        /// <summary>
        /// Clones a glyph and remaps all composite component glyph IDs using the old→new mapping.
        /// Preserves bounding box, instructions, and all flags.
        /// </summary>
        /// <param name="original">The original glyph from the source font</param>
        /// <param name="oldToNew">Mapping from old glyph ID to new subset glyph ID</param>
        /// <returns>A new glyph ready for the subset font</returns>
        private Glyph CloneGlyphWithRemappedComponents(Glyph original, Dictionary<ushort, ushort> oldToNew)
        {
            // Fix inverted bounding boxes (common in .notdef and some broken fonts)
            short xMin = original.Header.xMin;
            short xMax = original.Header.xMax;
            short yMin = original.Header.yMin;
            short yMax = original.Header.yMax;

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

            GlyphHeader header = new GlyphHeader(original.Header.numberOfContours,
                new BoundingRectangle(xMin, yMin, xMax, yMax));

            // Simple glyph
            if (original.Header.numberOfContours > 0 && original.SimpleData != null)
            {
                SimpleGlyph simple = new SimpleGlyph
                {
                    EndPtsOfContours = (ushort[])original.SimpleData.EndPtsOfContours.Clone(),
                    Instructions = (byte[])original.SimpleData.Instructions.Clone(),
                    XBytes = (byte[])original.SimpleData.XBytes.Clone(),
                    YBytes = (byte[])original.SimpleData.YBytes.Clone(),
                    FlagRuns = new List<FlagRun>()
                };

                foreach (FlagRun run in original.SimpleData.FlagRuns)
                {
                    simple.FlagRuns.Add(new FlagRun { Flag = run.Flag, RepeatCount = run.RepeatCount });
                }

                return new Glyph { Header = header, SimpleData = simple };
            }

            // Composite glyph – remap all components
            if (original.Header.numberOfContours < 0 && original.CompositeData != null)
            {
                CompositeGlyph composite = new CompositeGlyph
                {
                    Instructions = (byte[])original.CompositeData.Instructions.Clone(),
                    Components = new List<GlyphComponent>()
                };

                foreach (GlyphComponent comp in original.CompositeData.Components)
                {
                    ushort newGid = 0;
                    if (!oldToNew.TryGetValue(comp.GlyphIndex, out newGid))
                    {
                        // This should never happen if ResolveCompositeGlyphs was called correctly
                        System.Diagnostics.Debug.WriteLine(
                            string.Format("WARNING: Component glyph GID {0} not found in subset. Using .notdef.", comp.GlyphIndex));
                        newGid = 0;
                    }

                    GlyphComponent newComp = new GlyphComponent
                    {
                        Flags = comp.Flags,
                        GlyphIndex = newGid,
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

            // Empty glyph (e.g. space)
            return new Glyph { Header = header };
        }

        private void SubsetOS2(OpenTypeFont newFont, OpenTypeFont originalFont, HashSet<uint> usedCodePoints)
        {
            Os2Table original = originalFont.Os2Table;
            if (original == null)
                return;

            Os2Table os2 = original.Clone();

            // Default and break characters – space is the only safe choice in a subset
            os2.usDefaultChar = 32; // space
            os2.usBreakChar = 32; // space

            // --------------------------------------------------------------------
            // Critical for Windows Font Viewer (fontview.exe) to display the font
            // Without reasonable values here, Windows refuses to show a preview
            // --------------------------------------------------------------------
            if (usedCodePoints.Any())
            {
                uint first = usedCodePoints.Min();
                uint last = usedCodePoints.Max();

                os2.usFirstCharIndex = (ushort)Math.Max(32, Math.Min(first, 0xFFFF));
                os2.usLastCharIndex = (ushort)Math.Min(last, 0xFFFF);
            }
            else
            {
                os2.usFirstCharIndex = 32;
                os2.usLastCharIndex = 32;
            }

            // --------------------------------------------------------------------
            // Recalculate usWinAscent / usWinDescent based on actual glyphs in subset
            // Prevents clipping in Windows applications (especially important for accented chars)
            // --------------------------------------------------------------------
            short maxAscender = short.MinValue;
            short minDescender = short.MaxValue;

            foreach (Glyph glyph in newFont.GlyfTable.Glyphs)
            {
                if (glyph?.Header != null)
                {
                    if (glyph.Header.yMax > maxAscender) maxAscender = glyph.Header.yMax;
                    if (glyph.Header.yMin < minDescender) minDescender = glyph.Header.yMin;
                }
            }

            // Only update if the subset actually exceeds original values
            if (maxAscender > 0 && maxAscender > os2.usWinAscent)
                os2.usWinAscent = (ushort)maxAscender;

            if (minDescender < 0 && (ushort)(-minDescender) > os2.usWinDescent)
                os2.usWinDescent = (ushort)(-minDescender);

            // --------------------------------------------------------------------
            // Keep hhea table in sync with OS/2 typo metrics (best practice)
            // --------------------------------------------------------------------
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

        private void SubsetKern(OpenTypeFont newFont, OpenTypeFont originalFont, Dictionary<ushort, ushort> oldToNewGlyphId)
        {
            var originalKern = originalFont.KernTable;
            if (originalKern == null || originalKern.SubTables.Count == 0)
                return;

            var newKern = new KernTable
            {
                version = originalKern.version,
                numberOfFormat0Tables = 0 // vi räknar senare
            };

            foreach (var originalSubTable in originalKern.SubTables)
            {
                if (originalSubTable.coverage.Format != 0 || originalSubTable.Format0Subtable == null)
                    continue; // bara format 0 stöd för nu

                var format0 = originalSubTable.Format0Subtable;
                var newPairs = new List<KerningPair>();

                foreach (var pair in format0.Pairs)
                {
                    // Kolla om båda glyferna finns i subset
                    if (oldToNewGlyphId.TryGetValue(pair.left, out ushort newLeft) &&
                        oldToNewGlyphId.TryGetValue(pair.right, out ushort newRight))
                    {
                        newPairs.Add(new KerningPair(null)
                        {
                            left = newLeft,
                            right = newRight,
                            value = pair.value,
                            Combined = ((uint)newLeft << 16) | newRight
                        });
                    }
                }

                if (newPairs.Count == 0)
                    continue;

                // Sortera för binärsökning (som specen kräver)
                newPairs.Sort((a, b) => a.Combined.CompareTo(b.Combined));

                var newSubTable = new KernSubTable
                {
                    version = originalSubTable.version,
                    coverage = originalSubTable.coverage,
                    Format0Subtable = new KernSubTableFormat0(null)
                    {
                        nPairs = (ushort)newPairs.Count,
                        Pairs = newPairs.ToArray()
                    }
                };

                newKern.SubTables.Add(newSubTable);
                newKern.numberOfFormat0Tables++;
            }

            if (newKern.SubTables.Count > 0)
                newFont.AddOrReplaceTable(newKern);
        }
    }
}