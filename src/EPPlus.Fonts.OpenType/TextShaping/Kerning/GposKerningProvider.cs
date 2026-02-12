using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.TextShaping.Kerning
{
    internal class GposKerningProvider
    {
        private readonly Dictionary<uint, short> _kerningIndex;

        public GposKerningProvider(GposTable gposTable)
        {
            var sw = Stopwatch.StartNew();
            _kerningIndex = BuildKerningIndex(gposTable);
            sw.Stop();
            Console.WriteLine($"GposKerningProvider initialization: {sw.ElapsedMilliseconds}ms, {_kerningIndex.Count} pairs");
        }

        public short GetKerning(ushort leftGlyph, ushort rightGlyph)
        {
            uint key = ((uint)leftGlyph << 16) | rightGlyph;

            if (_kerningIndex.TryGetValue(key, out short kerning))
                return kerning;

            return 0;
        }

        private Dictionary<uint, short> BuildKerningIndex(GposTable gpos)
        {
            var index = new Dictionary<uint, short>();

            if (gpos == null)
                return index;

            var sw = Stopwatch.StartNew();
            var subtables = FindAllKerningSubtables(gpos);
            Console.WriteLine($"  Found {subtables.Count} kerning subtables in {sw.ElapsedMilliseconds}ms");

            int format1Count = 0;
            int format2Count = 0;

            foreach (var subtable in subtables)
            {
                sw.Stop();
                sw.Reset();
                sw.Start();

                if (subtable is PairPosSubTableFormat1 format1)
                {
                    IndexFormat1Subtable(format1, index);
                    format1Count++;
                    Console.WriteLine($"  Format1 subtable indexed: {sw.ElapsedMilliseconds}ms");
                }
                else if (subtable is PairPosSubTableFormat2 format2)
                {
                    IndexFormat2Subtable(format2, index);
                    format2Count++;
                    Console.WriteLine($"  Format2 subtable indexed: {sw.ElapsedMilliseconds}ms");
                }
            }

            Console.WriteLine($"  Total: {format1Count} Format1, {format2Count} Format2 subtables");
            Console.WriteLine($"  Final index size: {index.Count} pairs");

            return index;
        }

        private void IndexFormat1Subtable(PairPosSubTableFormat1 subtable, Dictionary<uint, short> index)
        {
            var coverage = subtable.Coverage;
            if (coverage == null || subtable.PairSets == null)
                return;

            for (ushort leftGlyph = 0; leftGlyph < 65535; leftGlyph++)
            {
                int coverageIndex = coverage.GetGlyphIndex(leftGlyph);
                if (coverageIndex < 0 || coverageIndex >= subtable.PairSets.Count)
                    continue;

                var pairSet = subtable.PairSets[coverageIndex];
                if (pairSet?.PairValueRecords == null)
                    continue;

                foreach (var pairRecord in pairSet.PairValueRecords)
                {
                    ushort rightGlyph = pairRecord.SecondGlyph;

                    short kerning = 0;
                    if (pairRecord.Value1?.XAdvance != 0)
                        kerning = pairRecord.Value1.XAdvance;

                    if (kerning != 0)
                    {
                        uint key = ((uint)leftGlyph << 16) | rightGlyph;
                        index[key] = kerning;
                    }
                }
            }
        }

        private void IndexFormat2Subtable(PairPosSubTableFormat2 subtable, Dictionary<uint, short> index)
        {
            if (subtable.Coverage == null || subtable.ClassMatrix == null)
                return;

            var classDef1 = subtable.ClassDef1;
            var classDef2 = subtable.ClassDef2;

            if (classDef1 == null || classDef2 == null)
                return;

            var sw = Stopwatch.StartNew();

            // ⚠️ BOTTLENECK: GetGlyphsFromCoverage kan vara långsam
            var leftGlyphs = GetGlyphsFromCoverage(subtable.Coverage);
            Console.WriteLine($"    GetGlyphsFromCoverage: {sw.ElapsedMilliseconds}ms, {leftGlyphs.Count} glyphs");

            sw.Stop();sw.Reset();sw.Start();
            int pairsAdded = 0;

            foreach (ushort leftGlyph in leftGlyphs)
            {
                int class1 = classDef1.GetClass(leftGlyph);
                if (class1 < 0 || class1 >= subtable.Class1Count)
                    continue;

                // ⚠️ BOTTLENECK: Loop through ALL 65535 possible right glyphs!
                for (ushort rightGlyph = 0; rightGlyph < 65535; rightGlyph++)
                {
                    int class2 = classDef2.GetClass(rightGlyph);
                    if (class2 < 0 || class2 >= subtable.Class2Count)
                        continue;

                    var record = subtable.ClassMatrix[class1, class2];
                    if (record == null)
                        continue;

                    short kerning = 0;
                    if (record.Value1?.XAdvance != 0)
                        kerning = record.Value1.XAdvance;

                    if (kerning != 0)
                    {
                        uint key = ((uint)leftGlyph << 16) | rightGlyph;
                        index[key] = kerning;
                        pairsAdded++;
                    }
                }
            }

            Console.WriteLine($"    Format2 expansion: {sw.ElapsedMilliseconds}ms, {pairsAdded} pairs added");
        }

        private List<ushort> GetGlyphsFromCoverage(CoverageTable coverage)
        {
            var glyphs = new List<ushort>();

            if (coverage is CoverageTableFormat1 format1)
            {
                if (format1.GlyphArray != null)
                    glyphs.AddRange(format1.GlyphArray);
            }
            else if (coverage is CoverageTableFormat2 format2)
            {
                if (format2.RangeRecords != null)
                {
                    foreach (var range in format2.RangeRecords)
                    {
                        for (ushort glyph = range.StartGlyphID; glyph <= range.EndGlyphID; glyph++)
                        {
                            glyphs.Add(glyph);
                        }
                    }
                }
            }

            return glyphs;
        }

        private List<PairPosSubTable> FindAllKerningSubtables(GposTable gpos)
        {
            var subtables = new List<PairPosSubTable>();

            if (gpos == null)
                return subtables;

            foreach (var featureRecord in gpos.FeatureList.FeatureRecords)
            {
                if (featureRecord.FeatureTag.Value == "kern")
                {
                    var feature = featureRecord.FeatureTable;

                    foreach (var lookupIndex in feature.LookupListIndices)
                    {
                        if (lookupIndex >= gpos.LookupList.Lookups.Count)
                            continue;

                        var lookup = gpos.LookupList.Lookups[lookupIndex];

                        if (lookup.LookupType == 2)
                        {
                            foreach (var subtable in lookup.SubTables)
                            {
                                if (subtable is PairPosSubTable pairPos)
                                {
                                    subtables.Add(pairPos);
                                }
                            }
                        }
                    }
                }
            }

            return subtables;
        }
    }
}