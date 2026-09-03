/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/13/2026         EPPlus Software AB           GPOS serialization tests (semantic validation)
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType1;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType2;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tests.Serialization
{
    /// <summary>
    /// Tests for GPOS table serialization with semantic validation.
    /// These tests verify that positioning data survives roundtrip serialization,
    /// without requiring byte-perfect output (Format 2 may be expanded to Format 1, etc.)
    /// </summary>
    [TestClass]
    public class GposSerializationTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        #region Helper Classes for .NET 3.5 Compatibility

        private class GlyphPair
        {
            public ushort FirstGlyph { get; set; }
            public ushort SecondGlyph { get; set; }

            public GlyphPair(ushort first, ushort second)
            {
                FirstGlyph = first;
                SecondGlyph = second;
            }

            public override bool Equals(object obj)
            {
                var other = obj as GlyphPair;
                if (other == null) return false;
                return FirstGlyph == other.FirstGlyph && SecondGlyph == other.SecondGlyph;
            }

            public override int GetHashCode()
            {
                return (FirstGlyph << 16) | SecondGlyph;
            }
        }

        private class AnchorPointPair
        {
            public short MarkAnchorX { get; set; }
            public short MarkAnchorY { get; set; }
            public short BaseAnchorX { get; set; }
            public short BaseAnchorY { get; set; }
        }

        private class SinglePosValue
        {
            public short XPlacement { get; set; }
            public short YPlacement { get; set; }
            public short XAdvance { get; set; }
            public short YAdvance { get; set; }
        }

        #endregion

        [TestMethod]
        public void Diagnose_SerializedFontOffsets()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            Debug.WriteLine("=== ORIGINAL TABLE RECORDS ===");
            foreach (var kvp in font.TableRecords)
            {
                Debug.WriteLine(string.Format("{0}: Offset={1}, Length={2}",
                    kvp.Key, kvp.Value.Offset, kvp.Value.Length));
            }

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();

            Debug.WriteLine(string.Format("\n=== SERIALIZED FONT: {0} bytes ===", bytes.Length));

            using (var ms = new MemoryStream(bytes))
            using (var reader = new FontsBinaryReader(ms))
            {
                reader.BaseStream.Position = 0;
                uint sfntVersion = reader.ReadUInt32BigEndian();
                ushort numTables = reader.ReadUInt16BigEndian();
                reader.ReadUInt16BigEndian();
                reader.ReadUInt16BigEndian();
                reader.ReadUInt16BigEndian();

                Debug.WriteLine(string.Format("\nsfntVersion: 0x{0:X8}", sfntVersion));
                Debug.WriteLine(string.Format("numTables: {0}", numTables));

                Debug.WriteLine("\n=== SERIALIZED TABLE RECORDS ===");
                for (int i = 0; i < numTables; i++)
                {
                    byte[] tagBytes = reader.ReadBytes(4);
                    string tag = System.Text.Encoding.ASCII.GetString(tagBytes);
                    uint checksum = reader.ReadUInt32BigEndian();
                    uint offset = reader.ReadUInt32BigEndian();
                    uint length = reader.ReadUInt32BigEndian();

                    string status = offset + length > bytes.Length ? " *** INVALID: extends beyond file!" : "";
                    Debug.WriteLine(string.Format("{0}: Offset={1}, Length={2}{3}", tag, offset, length, status));
                }
            }
        }

        #region Structure Preservation Tests

        [TestMethod]
        public void SerializeGpos_StructurePreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");
            var originalGpos = font.GposTable;

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            Assert.IsNotNull(reparsed.GposTable, "GPOS table should exist after roundtrip");
            Assert.AreEqual(originalGpos.MajorVersion, reparsed.GposTable.MajorVersion);
            Assert.AreEqual(originalGpos.MinorVersion, reparsed.GposTable.MinorVersion);

            Assert.AreEqual(
                originalGpos.ScriptList.ScriptRecords.Count,
                reparsed.GposTable.ScriptList.ScriptRecords.Count,
                "Script count should match");

            Assert.AreEqual(
                originalGpos.FeatureList.FeatureRecords.Count,
                reparsed.GposTable.FeatureList.FeatureRecords.Count,
                "Feature count should match");

            Assert.AreEqual(
                originalGpos.LookupList.Lookups.Count,
                reparsed.GposTable.LookupList.Lookups.Count,
                "Lookup count should match");

            for (int i = 0; i < originalGpos.LookupList.Lookups.Count; i++)
            {
                Assert.AreEqual(
                    originalGpos.LookupList.Lookups[i].LookupType,
                    reparsed.GposTable.LookupList.Lookups[i].LookupType,
                    string.Format("Lookup[{0}] type should match", i));
            }
        }

        [TestMethod]
        public void SerializeGpos_FeatureTagsPreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var originalTags = new List<string>();
            foreach (var feature in font.GposTable.FeatureList.FeatureRecords)
                originalTags.Add(feature.FeatureTag.Value);
            originalTags.Sort();

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var reparsedTags = new List<string>();
            foreach (var feature in reparsed.GposTable.FeatureList.FeatureRecords)
                reparsedTags.Add(feature.FeatureTag.Value);
            reparsedTags.Sort();

            Assert.AreEqual(originalTags.Count, reparsedTags.Count, "Feature count should match");
            for (int i = 0; i < originalTags.Count; i++)
            {
                Assert.AreEqual(originalTags[i], reparsedTags[i],
                    string.Format("Feature tag[{0}] should match", i));
            }

            Debug.WriteLine(string.Format("Features preserved: {0}", string.Join(", ", reparsedTags.ToArray())));
        }

        [TestMethod]
        public void SerializeGpos_ScriptTagsPreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var originalTags = new List<string>();
            foreach (var script in font.GposTable.ScriptList.ScriptRecords)
                originalTags.Add(script.ScriptTag.Value);
            originalTags.Sort();

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var reparsedTags = new List<string>();
            foreach (var script in reparsed.GposTable.ScriptList.ScriptRecords)
                reparsedTags.Add(script.ScriptTag.Value);
            reparsedTags.Sort();

            Assert.AreEqual(originalTags.Count, reparsedTags.Count, "Script count should match");
            for (int i = 0; i < originalTags.Count; i++)
            {
                Assert.AreEqual(originalTags[i], reparsedTags[i],
                    string.Format("Script tag[{0}] should match", i));
            }

            Debug.WriteLine(string.Format("Scripts preserved: {0}", string.Join(", ", reparsedTags.ToArray())));
        }

        #endregion

        #region PairPos (Type 2) Kerning Tests

        [TestMethod]
        public void SerializeGpos_PairPos_KerningValuesPreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var originalKerning = CollectKerningPairs(font);
            Debug.WriteLine(string.Format("Original font has {0} kerning pairs", originalKerning.Count));
            Assert.IsTrue(originalKerning.Count > 0, "Font should have kerning pairs");

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var reparsedKerning = CollectKerningPairs(reparsed);

            Assert.AreEqual(originalKerning.Count, reparsedKerning.Count, "Kerning pair count should match");

            int verified = 0;
            foreach (var pair in originalKerning.Keys)
            {
                short expectedValue = originalKerning[pair];

                Assert.IsTrue(reparsedKerning.ContainsKey(pair),
                    string.Format("Missing kerning pair ({0}, {1})", pair.FirstGlyph, pair.SecondGlyph));

                short actualValue = reparsedKerning[pair];
                Assert.AreEqual(expectedValue, actualValue,
                    string.Format("Kerning value mismatch for ({0}, {1}): expected {2}, got {3}",
                        pair.FirstGlyph, pair.SecondGlyph, expectedValue, actualValue));

                verified++;
            }

            Debug.WriteLine(string.Format("Verified {0} kerning pairs", verified));
        }

        [TestMethod]
        public void SerializeGpos_PairPos_SpecificPairsVerified()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            ushort fGlyph, eGlyph, aGlyph, vGlyph;
            font.CmapTable.TryGetGlyphId('f', out fGlyph);
            font.CmapTable.TryGetGlyphId('e', out eGlyph);
            font.CmapTable.TryGetGlyphId('A', out aGlyph);
            font.CmapTable.TryGetGlyphId('V', out vGlyph);

            var origLookup = FindFirstLookupOfType(font.GposTable, 2);
            Assert.IsNotNull(origLookup, "Should have PairPos lookup");

            var origSubtable = origLookup.SubTables[0] as PairPosSubTableFormat1;
            Assert.IsNotNull(origSubtable, "Should have PairPos Format 1 subtable");

            ValueRecord feOrig1, feOrig2, avOrig1, avOrig2;
            bool hasFe = origSubtable.TryGetPairAdjustment(fGlyph, eGlyph, out feOrig1, out feOrig2);
            bool hasAv = origSubtable.TryGetPairAdjustment(aGlyph, vGlyph, out avOrig1, out avOrig2);

            if (hasFe) Debug.WriteLine(string.Format("Original: f-e = {0}", feOrig1.XAdvance));
            if (hasAv) Debug.WriteLine(string.Format("Original: A-V = {0}", avOrig1.XAdvance));

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var newLookup = FindFirstLookupOfType(reparsed.GposTable, 2);
            Assert.IsNotNull(newLookup, "Reparsed should have PairPos lookup");

            var newSubtable = newLookup.SubTables[0] as PairPosSubTableFormat1;
            Assert.IsNotNull(newSubtable, "Reparsed should have PairPos Format 1 subtable");

            if (hasFe)
            {
                ValueRecord feNew1, feNew2;
                Assert.IsTrue(newSubtable.TryGetPairAdjustment(fGlyph, eGlyph, out feNew1, out feNew2), "f-e pair should exist");
                Assert.AreEqual(feOrig1.XAdvance, feNew1.XAdvance, "f-e kerning should match");
                Debug.WriteLine(string.Format("f-e: {0} verified", feNew1.XAdvance));
            }

            if (hasAv)
            {
                ValueRecord avNew1, avNew2;
                Assert.IsTrue(newSubtable.TryGetPairAdjustment(aGlyph, vGlyph, out avNew1, out avNew2), "A-V pair should exist");
                Assert.AreEqual(avOrig1.XAdvance, avNew1.XAdvance, "A-V kerning should match");
                Debug.WriteLine(string.Format("A-V: {0} verified", avNew1.XAdvance));
            }
        }

        [TestMethod]
        public void SerializeGpos_PairPos_ValueFormatPreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var origLookup = FindFirstLookupOfType(font.GposTable, 2);
            var origSubtable = origLookup.SubTables[0] as PairPosSubTableFormat1;

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var newLookup = FindFirstLookupOfType(reparsed.GposTable, 2);
            var newSubtable = newLookup.SubTables[0] as PairPosSubTableFormat1;

            Assert.AreEqual(origSubtable.ValueFormat1, newSubtable.ValueFormat1, "ValueFormat1 should be preserved");
            Assert.AreEqual(origSubtable.ValueFormat2, newSubtable.ValueFormat2, "ValueFormat2 should be preserved");

            Debug.WriteLine(string.Format("ValueFormat1: 0x{0:X4}", newSubtable.ValueFormat1));
            Debug.WriteLine(string.Format("ValueFormat2: 0x{0:X4}", newSubtable.ValueFormat2));
        }

        #endregion

        #region SinglePos (Type 1) Tests

        [TestMethod]
        public void SerializeGpos_SinglePos_ValuesPreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var singlePosLookup = FindFirstLookupOfType(font.GposTable, 1);
            if (singlePosLookup == null)
            {
                Assert.Inconclusive("Roboto does not have SinglePos lookups");
                return;
            }

            var originalAdjustments = CollectSinglePosAdjustments(font);
            Debug.WriteLine(string.Format("Original has {0} single adjustments", originalAdjustments.Count));

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var reparsedAdjustments = CollectSinglePosAdjustments(reparsed);

            Assert.AreEqual(originalAdjustments.Count, reparsedAdjustments.Count, "SinglePos adjustment count should match");

            foreach (ushort glyphId in originalAdjustments.Keys)
            {
                var expected = originalAdjustments[glyphId];

                Assert.IsTrue(reparsedAdjustments.ContainsKey(glyphId),
                    string.Format("Missing adjustment for glyph {0}", glyphId));

                var actual = reparsedAdjustments[glyphId];

                Assert.AreEqual(expected.XPlacement, actual.XPlacement, string.Format("Glyph {0} XPlacement", glyphId));
                Assert.AreEqual(expected.YPlacement, actual.YPlacement, string.Format("Glyph {0} YPlacement", glyphId));
                Assert.AreEqual(expected.XAdvance, actual.XAdvance, string.Format("Glyph {0} XAdvance", glyphId));
                Assert.AreEqual(expected.YAdvance, actual.YAdvance, string.Format("Glyph {0} YAdvance", glyphId));
            }

            Debug.WriteLine(string.Format("Verified {0} single adjustments", originalAdjustments.Count));
        }

        #endregion

        #region MarkToBase (Type 4) Tests

        [TestMethod]
        public void SerializeGpos_MarkToBase_StructurePreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var markToBaseLookup = FindFirstLookupOfType(font.GposTable, 4);
            if (markToBaseLookup == null)
            {
                Assert.Inconclusive("Roboto does not have MarkToBase lookups");
                return;
            }

            var origSubtable = markToBaseLookup.SubTables[0] as MarkToBaseSubTableFormat1;
            Assert.IsNotNull(origSubtable, "Should have MarkToBase subtable");

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var newLookup = FindFirstLookupOfType(reparsed.GposTable, 4);
            Assert.IsNotNull(newLookup, "MarkToBase lookup should exist");

            var newSubtable = newLookup.SubTables[0] as MarkToBaseSubTableFormat1;
            Assert.IsNotNull(newSubtable, "MarkToBase subtable should exist");

            Assert.AreEqual(origSubtable.MarkClassCount, newSubtable.MarkClassCount, "MarkClassCount");
            Assert.AreEqual(origSubtable.MarkArray.MarkCount, newSubtable.MarkArray.MarkCount, "MarkCount");
            Assert.AreEqual(origSubtable.BaseArray.BaseCount, newSubtable.BaseArray.BaseCount, "BaseCount");

            Debug.WriteLine(string.Format("MarkToBase preserved: {0} classes, {1} marks, {2} bases",
                newSubtable.MarkClassCount, newSubtable.MarkArray.MarkCount, newSubtable.BaseArray.BaseCount));
        }

        [TestMethod]
        public void SerializeGpos_MarkToBase_AnchorPointsPreserved()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var markToBaseLookup = FindFirstLookupOfType(font.GposTable, 4);
            if (markToBaseLookup == null)
            {
                Assert.Inconclusive("Roboto does not have MarkToBase lookups");
                return;
            }

            var origSubtable = markToBaseLookup.SubTables[0] as MarkToBaseSubTableFormat1;
            var originalAttachments = CollectMarkToBaseAttachments(origSubtable);

            if (originalAttachments.Count == 0)
            {
                Assert.Inconclusive("No attachments found");
                return;
            }

            Debug.WriteLine(string.Format("Found {0} mark-base attachments", originalAttachments.Count));

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            var newLookup = FindFirstLookupOfType(reparsed.GposTable, 4);
            var newSubtable = newLookup.SubTables[0] as MarkToBaseSubTableFormat1;

            int verified = 0;
            int maxToVerify = 50;

            foreach (var pair in originalAttachments.Keys)
            {
                if (verified >= maxToVerify) break;

                var expected = originalAttachments[pair];

                AnchorTable newMarkAnchor, newBaseAnchor;
                Assert.IsTrue(newSubtable.TryGetAttachment(pair.FirstGlyph, pair.SecondGlyph, out newMarkAnchor, out newBaseAnchor),
                    string.Format("Missing attachment for mark {0}, base {1}", pair.FirstGlyph, pair.SecondGlyph));

                Assert.AreEqual(expected.MarkAnchorX, newMarkAnchor.XCoordinate,
                    string.Format("Mark anchor X for ({0}, {1})", pair.FirstGlyph, pair.SecondGlyph));
                Assert.AreEqual(expected.MarkAnchorY, newMarkAnchor.YCoordinate,
                    string.Format("Mark anchor Y for ({0}, {1})", pair.FirstGlyph, pair.SecondGlyph));
                Assert.AreEqual(expected.BaseAnchorX, newBaseAnchor.XCoordinate,
                    string.Format("Base anchor X for ({0}, {1})", pair.FirstGlyph, pair.SecondGlyph));
                Assert.AreEqual(expected.BaseAnchorY, newBaseAnchor.YCoordinate,
                    string.Format("Base anchor Y for ({0}, {1})", pair.FirstGlyph, pair.SecondGlyph));

                verified++;
            }

            Debug.WriteLine(string.Format("Verified {0} anchor point pairs", verified));
        }

        #endregion

        #region Feature-Lookup Index Integrity Tests

        [TestMethod]
        public void SerializeGpos_FeatureLookupIndices_AreValid()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            int lookupCount = reparsed.GposTable.LookupList.Lookups.Count;

            foreach (var feature in reparsed.GposTable.FeatureList.FeatureRecords)
            {
                foreach (var idx in feature.FeatureTable.LookupListIndices)
                {
                    Assert.IsTrue(idx < lookupCount,
                        string.Format("Feature '{0}' references invalid lookup index {1} (max={2})",
                            feature.FeatureTag.Value, idx, lookupCount - 1));
                }
            }

            Debug.WriteLine(string.Format("All feature->lookup indices valid (max index: {0})", lookupCount - 1));
        }

        [TestMethod]
        public void Diagnose_GposTableOffset()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            Debug.WriteLine("=== TABLE RECORDS ===");
            foreach (var kvp in font.TableRecords)
            {
                Debug.WriteLine(string.Format("{0}: Offset={1}, Length={2}",
                    kvp.Key, kvp.Value.Offset, kvp.Value.Length));
            }

            Debug.WriteLine(string.Format("\n=== READER STATE ==="));
            Debug.WriteLine(string.Format("Reader stream length: {0}", font._tblSettings.TableReaderFactory.FontBytesLength));

            Debug.WriteLine("\n=== LOADING GPOS ===");
            var gpos = font.GposTable;
            Debug.WriteLine(string.Format("GPOS version: {0}.{1}", gpos.MajorVersion, gpos.MinorVersion));
        }

        [TestMethod]
        public void SerializeGpos_LangSysFeatureIndices_AreValid()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            int featureCount = reparsed.GposTable.FeatureList.FeatureRecords.Count;

            foreach (var script in reparsed.GposTable.ScriptList.ScriptRecords)
            {
                if (script.ScriptTable.DefaultLangSys != null)
                {
                    foreach (var idx in script.ScriptTable.DefaultLangSys.FeatureIndices)
                    {
                        Assert.IsTrue(idx < featureCount,
                            string.Format("Script '{0}' DefaultLangSys references invalid feature index {1}",
                                script.ScriptTag.Value, idx));
                    }
                }

                foreach (var langSys in script.ScriptTable.LangSysRecords)
                {
                    foreach (var idx in langSys.LangSysTable.FeatureIndices)
                    {
                        Assert.IsTrue(idx < featureCount,
                            string.Format("Script '{0}' LangSys '{1}' references invalid feature index {2}",
                                script.ScriptTag.Value, langSys.LangSysTag, idx));
                    }
                }
            }

            Debug.WriteLine(string.Format("All LangSys->feature indices valid (max index: {0})", featureCount - 1));
        }

        #endregion

        #region Coverage Table Integrity Tests

        [TestMethod]
        public void SerializeGpos_CoverageGlyphIds_AreValid()
        {
            var font = TestFolderEngine.LoadFont("Roboto");

            var serializer = new OpenTypeFontSerializer(font);
            var bytes = serializer.Serialize();
            var reparsed = new OpenTypeFont(bytes);

            ushort maxGlyph = reparsed.MaxpTable.numGlyphs;

            for (int lookupIdx = 0; lookupIdx < reparsed.GposTable.LookupList.Lookups.Count; lookupIdx++)
            {
                var lookup = reparsed.GposTable.LookupList.Lookups[lookupIdx];

                for (int subtableIdx = 0; subtableIdx < lookup.SubTables.Count; subtableIdx++)
                {
                    var subtable = lookup.SubTables[subtableIdx];
                    var coverages = GetCoveragesFromSubtable(subtable);

                    foreach (var coverage in coverages)
                    {
                        var glyphIds = coverage.GetCoveredGlyphs();
                        foreach (var glyphId in glyphIds)
                        {
                            Assert.IsTrue(glyphId < maxGlyph,
                                string.Format("Lookup type {0} subtable {1}: Coverage references invalid glyph {2} (max={3})",
                                    lookup.LookupType, subtableIdx, glyphId, maxGlyph - 1));
                        }
                    }
                }
            }

            Debug.WriteLine(string.Format("All coverage glyph IDs valid (max glyph: {0})", maxGlyph - 1));
        }

        #endregion

        #region Helper Methods

        private LookupTable FindFirstLookupOfType(GposTable gpos, int lookupType)
        {
            if (gpos == null) return null;

            foreach (var lookup in gpos.LookupList.Lookups)
            {
                if (lookup.LookupType == lookupType)
                    return lookup;
            }
            return null;
        }

        private Dictionary<GlyphPair, short> CollectKerningPairs(OpenTypeFont font)
        {
            var result = new Dictionary<GlyphPair, short>();

            var pairPosLookup = FindFirstLookupOfType(font.GposTable, 2);
            if (pairPosLookup == null) return result;

            foreach (var subtableObj in pairPosLookup.SubTables)
            {
                var subtable = subtableObj as PairPosSubTableFormat1;
                if (subtable == null) continue;

                var coveredGlyphs = subtable.Coverage.GetCoveredGlyphs();

                for (int i = 0; i < coveredGlyphs.Length && i < subtable.PairSets.Count; i++)
                {
                    ushort firstGlyph = coveredGlyphs[i];
                    var pairSet = subtable.PairSets[i];

                    if (pairSet == null) continue;

                    foreach (var pair in pairSet.PairValueRecords)
                    {
                        var key = new GlyphPair(firstGlyph, pair.SecondGlyph);
                        result[key] = pair.Value1.XAdvance;
                    }
                }
            }

            return result;
        }

        private Dictionary<ushort, SinglePosValue> CollectSinglePosAdjustments(OpenTypeFont font)
        {
            var result = new Dictionary<ushort, SinglePosValue>();

            var singlePosLookup = FindFirstLookupOfType(font.GposTable, 1);
            if (singlePosLookup == null) return result;

            foreach (var subtableObj in singlePosLookup.SubTables)
            {
                var f1 = subtableObj as SinglePosSubTableFormat1;
                if (f1 != null)
                {
                    var glyphIds = f1.Coverage.GetCoveredGlyphs();
                    foreach (var glyphId in glyphIds)
                    {
                        result[glyphId] = new SinglePosValue
                        {
                            XPlacement = f1.Value.XPlacement,
                            YPlacement = f1.Value.YPlacement,
                            XAdvance = f1.Value.XAdvance,
                            YAdvance = f1.Value.YAdvance
                        };
                    }
                    continue;
                }

                var f2 = subtableObj as SinglePosSubTableFormat2;
                if (f2 != null)
                {
                    var glyphIds = f2.Coverage.GetCoveredGlyphs();
                    for (int i = 0; i < glyphIds.Length && i < f2.Values.Length; i++)
                    {
                        result[glyphIds[i]] = new SinglePosValue
                        {
                            XPlacement = f2.Values[i].XPlacement,
                            YPlacement = f2.Values[i].YPlacement,
                            XAdvance = f2.Values[i].XAdvance,
                            YAdvance = f2.Values[i].YAdvance
                        };
                    }
                }
            }

            return result;
        }

        private Dictionary<GlyphPair, AnchorPointPair> CollectMarkToBaseAttachments(
            MarkToBaseSubTableFormat1 subtable)
        {
            var result = new Dictionary<GlyphPair, AnchorPointPair>();

            var markGlyphs = subtable.MarkCoverage.GetCoveredGlyphs();
            var baseGlyphs = subtable.BaseCoverage.GetCoveredGlyphs();

            int markLimit = markGlyphs.Length > 100 ? 100 : markGlyphs.Length;
            int baseLimit = baseGlyphs.Length > 100 ? 100 : baseGlyphs.Length;

            for (int m = 0; m < markLimit; m++)
            {
                ushort markGlyph = markGlyphs[m];

                for (int b = 0; b < baseLimit; b++)
                {
                    ushort baseGlyph = baseGlyphs[b];

                    AnchorTable markAnchor, baseAnchor;
                    if (subtable.TryGetAttachment(markGlyph, baseGlyph, out markAnchor, out baseAnchor))
                    {
                        var key = new GlyphPair(markGlyph, baseGlyph);
                        result[key] = new AnchorPointPair
                        {
                            MarkAnchorX = markAnchor.XCoordinate,
                            MarkAnchorY = markAnchor.YCoordinate,
                            BaseAnchorX = baseAnchor.XCoordinate,
                            BaseAnchorY = baseAnchor.YCoordinate
                        };
                    }
                }
            }

            return result;
        }

        private List<CoverageTable> GetCoveragesFromSubtable(object subtable)
        {
            var coverages = new List<CoverageTable>();

            var pp1 = subtable as PairPosSubTableFormat1;
            if (pp1 != null) { coverages.Add(pp1.Coverage); return coverages; }

            var sp1 = subtable as SinglePosSubTableFormat1;
            if (sp1 != null) { coverages.Add(sp1.Coverage); return coverages; }

            var sp2 = subtable as SinglePosSubTableFormat2;
            if (sp2 != null) { coverages.Add(sp2.Coverage); return coverages; }

            var mtb = subtable as MarkToBaseSubTableFormat1;
            if (mtb != null) { coverages.Add(mtb.MarkCoverage); coverages.Add(mtb.BaseCoverage); return coverages; }

            return coverages;
        }

        #endregion
    }
}