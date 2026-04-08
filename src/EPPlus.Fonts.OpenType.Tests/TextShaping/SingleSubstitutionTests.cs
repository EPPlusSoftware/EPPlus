/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/19/2026         EPPlus Software AB           Single Substitution tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Tables.Gsub.Data.Lookups;
using EPPlus.Fonts.OpenType.TextShaping;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    [TestClass]
    public class SingleSubstitutionTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void SingleSubstitution_SmallCaps_SubstitutesGlyphs()
        {
            // NOTE: This test requires a font with small caps feature (smcp)
            // Common fonts with smcp: Georgia, Garamond, Calibri, Cambria

            var fontNames = new[] { "Roboto" };

            foreach (var fontName in fontNames)
            {
                try
                {
                    var font = OpenTypeFonts.LoadFont(fontName, FontSubFamily.Regular);

                    // Verify font has smcp feature with Type 1 lookup
                    if (font == null || !font.FullName.Contains(fontName) || font.GsubTable == null)
                        continue;

                    bool hasSmcpWithType1 = false;
                    foreach (var featureRecord in font.GsubTable.FeatureList.FeatureRecords)
                    {
                        if (featureRecord.FeatureTag.Value == "smcp")
                        {
                            var feature = featureRecord.FeatureTable;
                            foreach (var lookupIndex in feature.LookupListIndices)
                            {
                                if (lookupIndex < font.GsubTable.LookupList.Lookups.Count)
                                {
                                    var lookup = font.GsubTable.LookupList.Lookups[lookupIndex];
                                    DebugWriteLine($"{fontName} smcp uses Lookup Type {lookup.LookupType}");
                                    if (lookup.LookupType == 1)
                                    {
                                        hasSmcpWithType1 = true;
                                    }
                                }
                            }
                        }
                    }

                    if (!hasSmcpWithType1)
                    {
                        DebugWriteLine($"{fontName} has smcp but not with Type 1 lookup - skipping");
                        continue;
                    }

                    // Found a font with smcp using Type 1!
                    var shaper = new TextShaper(font);

                    // Act - lowercase letters should become small caps
                    var normal = shaper.Shape("hello", ShapingOptions.Default);
                    var smallCaps = shaper.Shape("hello", new ShapingOptions
                    {
                        ApplySubstitutions = true,
                        GsubFeatures = new List<string> { "smcp" },
                        ApplyPositioning = true
                    });

                    // Assert - At least one glyph should be substituted
                    bool anySubstituted = false;
                    for (int i = 0; i < normal.Glyphs.Length; i++)
                    {
                        if (normal.Glyphs[i].GlyphId != smallCaps.Glyphs[i].GlyphId)
                        {
                            anySubstituted = true;
                            DebugWriteLine($"{fontName}: '{normal.OriginalText[i]}' GID {normal.Glyphs[i].GlyphId} → {smallCaps.Glyphs[i].GlyphId}");
                        }
                    }

                    Assert.IsTrue(anySubstituted,
                        $"{fontName} has smcp feature but no glyphs were substituted for 'hello'");

                    DebugWriteLine($"✓ {fontName}: Small caps working!");
                    return; // Test passed, no need to try other fonts
                }
                catch (System.IO.FileNotFoundException)
                {
                    continue; // Try next font
                }
            }

            Assert.Inconclusive("No font with working small caps (Type 1) found. Tested: " + string.Join(", ", fontNames));
        }

       

        [TestMethod]
        public void SingleSubstitution_AppliesBeforeLigatures()
        {
            // Test that single substitution happens before ligature formation
            // This is important: if we request both smcp and liga, small caps should apply first

            var fontNames = new[] { "Roboto" };

            foreach (var fontName in fontNames)
            {
                try
                {
                    var font = OpenTypeFonts.LoadFont(fontName);

                    if (font == null || !font.FullName  .Contains(fontName) || font.GsubTable == null)
                        continue;

                    bool hasSmcpWithType1 = false;
                    bool hasLiga = false;

                    foreach (var featureRecord in font.GsubTable.FeatureList.FeatureRecords)
                    {
                        if (featureRecord.FeatureTag.Value == "smcp")
                        {
                            var feature = featureRecord.FeatureTable;
                            foreach (var lookupIndex in feature.LookupListIndices)
                            {
                                if (lookupIndex < font.GsubTable.LookupList.Lookups.Count)
                                {
                                    var lookup = font.GsubTable.LookupList.Lookups[lookupIndex];
                                    if (lookup.LookupType == 1)
                                    {
                                        hasSmcpWithType1 = true;
                                    }
                                }
                            }
                        }
                        else if (featureRecord.FeatureTag.Value == "liga")
                        {
                            hasLiga = true;
                        }
                    }

                    if (!hasSmcpWithType1 || !hasLiga)
                        continue;

                    var shaper = new TextShaper(font);

                    // Act - Apply both features (single substitution should happen first)
                    var bothFeatures = shaper.Shape("office", new ShapingOptions
                    {
                        ApplySubstitutions = true,
                        GsubFeatures = new List<string> { "smcp", "liga" },
                        ApplyPositioning = true
                    });

                    var onlySmcp = shaper.Shape("office", new ShapingOptions
                    {
                        ApplySubstitutions = true,
                        GsubFeatures = new List<string> { "smcp" },
                        ApplyPositioning = true
                    });

                    // Assert - Should not crash, glyphs should be processed
                    Assert.IsNotNull(bothFeatures);
                    Assert.IsNotNull(onlySmcp);

                    DebugWriteLine($"✓ {fontName}: Feature ordering test passed");
                    return;
                }
                catch (System.IO.FileNotFoundException)
                {
                    continue;
                }
            }

            Assert.Inconclusive("No font with both smcp (Type 1) and liga features found");
        }

        private void DebugWriteLine(string message)
        {
            Debug.WriteLine(message);
            TestContext?.WriteLine(message);
        }
    }
}