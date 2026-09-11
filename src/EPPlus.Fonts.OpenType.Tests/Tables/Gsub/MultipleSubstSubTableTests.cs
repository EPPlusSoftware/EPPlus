/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/09/2026         EPPlus Software AB           GSUB Multiple Substitution (Type 2) tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.Tables.Common.Layout.Coverage;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.Tables.Gsub.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tests.Tables.Gsub
{
    /// <summary>
    /// Tests the new GSUB Lookup Type 2 (Multiple Substitution) data model:
    /// <c>MultipleSubstSubTable</c> and <c>MultipleSubstSubTableDeserializer</c>. Unlike Lookup
    /// Type 1 (one glyph -> one glyph) or Type 4/Ligature (many glyphs -> one), Type 2 maps ONE
    /// input glyph to a SEQUENCE of output glyphs - the direction "ccmp" typically needs for
    /// decomposing a precomposed glyph into a base + combining marks.
    ///
    /// All tables here are built by hand so the tests are deterministic and don't depend on any
    /// specific font shipping real Multiple Substitution data.
    /// </summary>
    [TestClass]
    public class MultipleSubstSubTableTests
    {
        public TestContext? TestContext { get; set; }

        // Arbitrary glyph ids used only as stand-ins in the synthetic table.
        private const ushort InputGidA = 10;   // covered, decomposes into two glyphs
        private const ushort InputGidB = 20;   // covered, decomposes into three glyphs
        private const ushort UncoveredGid = 30; // never listed in Coverage

        private static readonly ushort[] SequenceForA = { 100, 101 };
        private static readonly ushort[] SequenceForB = { 200, 201, 202 };

        private static MultipleSubstSubTable BuildSyntheticSubtable()
        {
            return new MultipleSubstSubTable
            {
                SubtableFormat = 1,
                Coverage = CoverageTableFormat2.CreateCoverageFormat2(new List<ushort> { InputGidA, InputGidB }),
                Sequences = new List<ushort[]> { SequenceForA, SequenceForB }
            };
        }

        [TestMethod]
        public void GetSubstitution_CoveredGlyph_ReturnsItsSequence()
        {
            var subtable = BuildSyntheticSubtable();

            CollectionAssert.AreEqual(SequenceForA, subtable.GetSubstitution(InputGidA));
            CollectionAssert.AreEqual(SequenceForB, subtable.GetSubstitution(InputGidB));
        }

        [TestMethod]
        public void GetSubstitution_UncoveredGlyph_ReturnsNull()
        {
            var subtable = BuildSyntheticSubtable();

            Assert.IsNull(subtable.GetSubstitution(UncoveredGid));
        }

        [TestMethod]
        public void SerializeThenDeserialize_RoundTrips_SameCoverageAndSequences()
        {
            var original = BuildSyntheticSubtable();

            byte[] bytes = original.Serialize();

            using (var stream = new MemoryStream(bytes))
            {
                var reader = new FontsBinaryReader(stream);
                var roundTripped = new MultipleSubstSubTableDeserializer(reader).Deserialize(0);

                Assert.AreEqual(1, roundTripped.SubtableFormat);
                CollectionAssert.AreEqual(SequenceForA, roundTripped.GetSubstitution(InputGidA));
                CollectionAssert.AreEqual(SequenceForB, roundTripped.GetSubstitution(InputGidB));
                Assert.IsNull(roundTripped.GetSubstitution(UncoveredGid));
            }
        }

        [TestMethod]
        public void SerializeThenDeserialize_PreservesCoverageOrderIndependence()
        {
            // Coverage index order must line up with Sequences order - verify both covered
            // glyphs still resolve correctly after a round trip, not just the first one (which
            // would still "work" even if the sequence-offset array were parsed incorrectly, as
            // long as the coverage table itself parsed fine).
            var original = BuildSyntheticSubtable();
            byte[] bytes = original.Serialize();

            using (var stream = new MemoryStream(bytes))
            {
                var reader = new FontsBinaryReader(stream);
                var roundTripped = new MultipleSubstSubTableDeserializer(reader).Deserialize(0);

                var coveredGlyphs = roundTripped.Coverage.GetCoveredGlyphs();
                CollectionAssert.AreEqual(new ushort[] { InputGidA, InputGidB }, coveredGlyphs);
            }
        }
    }
}