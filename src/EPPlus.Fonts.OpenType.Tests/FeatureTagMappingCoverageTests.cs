/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/07/2026         EPPlus Software AB           Feature-tag mapping coverage
 *************************************************************************************************/
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests
{
    /// <summary>
    /// GsubFeatureTags.ToTagList and GposFeatureTags.ToTagList use an explicit if-chain rather
    /// than reflecting over the enum, specifically so a future named combination value (e.g. an
    /// "All = Kern | Mark" convenience member) cannot silently slip an extra tag into the list -
    /// see the discussion in Åtgärdsplan about GposFeature.All matching via bitwise AND.
    ///
    /// The trade-off is that the if-chain has to be updated by hand whenever a new single-bit
    /// flag is added to either enum. These tests are the safety net for that: they walk every
    /// enum member via reflection and require each single-bit flag to produce exactly one, and a
    /// DIFFERENT, tag. Forgetting to add a branch in ToTagList turns into a failing test here,
    /// rather than a silently dropped feature.
    /// </summary>
    [TestClass]
    public class FeatureTagMappingCoverageTests
    {
        public TestContext? TestContext { get; set; }

        [TestMethod]
        public void GsubFeatureTags_CoversEverySingleBitFlag()
        {
            AssertEveryFlagMapsToExactlyOneDistinctTag(
                GetSingleBitFlags<GsubFeature>(),
                flag => GsubFeatureTags.ToTagList((GsubFeature)flag));
        }

        [TestMethod]
        public void GposFeatureTags_CoversEverySingleBitFlag()
        {
            AssertEveryFlagMapsToExactlyOneDistinctTag(
                GetSingleBitFlags<GposFeature>(),
                flag => GposFeatureTags.ToTagList((GposFeature)flag));
        }

        /// <summary>
        /// Returns every defined enum member that is a single-bit flag (excludes None = 0 and any
        /// named combination such as a hypothetical "All" whose value has more than one bit set).
        /// A named combination is intentionally NOT tested here: its whole point is to be a
        /// convenience alias, not a distinct feature with its own tag, so it must not appear in
        /// ToTagList's if-chain at all.
        /// </summary>
        private static List<int> GetSingleBitFlags<TEnum>() where TEnum : struct, Enum
        {
            return Enum.GetValues(typeof(TEnum))
                .Cast<TEnum>()
                .Select(v => Convert.ToInt32(v))
                .Where(v => v != 0 && (v & (v - 1)) == 0) // v != 0 and exactly one bit set
                .Distinct()
                .ToList();
        }

        private static void AssertEveryFlagMapsToExactlyOneDistinctTag(
            List<int> singleBitFlags, Func<int, List<string>> toTagList)
        {
            Assert.AreNotEqual(
                0,
                singleBitFlags.Count,
                "the enum under test must actually have single-bit flags for this test to mean anything");

            var seenTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var flag in singleBitFlags)
            {
                var tags = toTagList(flag);

                Assert.AreEqual(
                    1,
                    tags.Count,
                    $"flag value {flag} (0x{flag:X}) must map to exactly one tag. Either it is "
                    + "missing a branch in ToTagList, or it is unexpectedly mapping to more than one.");

                Assert.IsTrue(
                    seenTags.Add(tags[0]),
                    $"tag \"{tags[0]}\" is produced by more than one flag - each single-bit flag "
                    + "must map to a distinct OpenType feature tag.");
            }
        }
    }
}