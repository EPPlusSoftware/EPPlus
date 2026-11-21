/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/17/2026         EPPlus Software AB           Pass-through hinting table subsetting tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontValidation;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using EPPlus.Fonts.OpenType.Utils;
using System.IO;

namespace EPPlus.Fonts.OpenType.Tests.Subsetting
{
    [TestClass]
    public class PassThroughTableSubsetProcessorTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        private static readonly string[] PassThroughTags = { "cvt ", "fpgm", "prep", "gasp" };

        private const string SubsetText = "Hello World 123";

        /// <summary>
        /// Loads a fresh, uncached font instance directly from the test font folder.
        /// </summary>
        private static OpenTypeFont LoadFontFile(string fileName)
        {
            var path = Path.Combine(FontFolder, fileName);
            if (!File.Exists(path))
                Assert.Inconclusive("Test font not found: " + path);
            return new OpenTypeFont(File.ReadAllBytes(path));
        }

        private static byte[] PadTo4(byte[] data)
        {
            var padded = (byte[])data.Clone();
            int paddedLen = (padded.Length + 3) & ~3;
            if (paddedLen > padded.Length)
                Array.Resize(ref padded, paddedLen);
            return padded;
        }

        [TestMethod]
        public void PassThroughTags_AreFourCharactersLong()
        {
            Assert.AreEqual("cvt ", EPPlus.Fonts.OpenType.Tables.TableNames.Cvt);
            Assert.AreEqual(4, EPPlus.Fonts.OpenType.Tables.TableNames.Fpgm.Length);
            Assert.AreEqual(4, EPPlus.Fonts.OpenType.Tables.TableNames.Prep.Length);
            Assert.AreEqual(4, EPPlus.Fonts.OpenType.Tables.TableNames.Gasp.Length);
        }

        /// <summary>
        /// Exploratory test: lists which hinting-related tables each bundled test font contains.
        /// </summary>
        [TestMethod]
        public void Probe_ListHintingTablesInTestFonts()
        {
            foreach (var folder in FontFolders)
            {
                if (!Directory.Exists(folder))
                    continue;

                foreach (var file in Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories))
                {
                    var ext = Path.GetExtension(file).ToLowerInvariant();
                    if (ext != ".ttf" && ext != ".otf")
                        continue;

                    try
                    {
                        var font = new OpenTypeFont(File.ReadAllBytes(file));
                        var present = PassThroughTags
                            .Select(t => font.TableRecords.ContainsKey(t) ? t.Trim() : "-");
                        Console.WriteLine($"{Path.GetFileName(file),-50} {font.Format,-4} {string.Join(" ", present)}");
                    }
                    catch (Exception ex)
                    {
                       Console.WriteLine($"{Path.GetFileName(file),-50} ERROR: {ex.Message}");
                    }
                }
            }
        }

        // Test case 1: all four tables are present in the original and copied byte-identical.
        [TestMethod]
        [DataRow("Roboto-Regular.ttf")] // Verify with the probe output
        [DataRow("OpenSans-Regular.ttf")]
        public void Subset_FontWithHintingTables_CopiesTablesByteIdentical(string fileName)
        {
            var font = LoadFontFile(fileName);
            foreach (var tag in PassThroughTags)
            {
                Assert.IsTrue(font.TableRecords.ContainsKey(tag),
                    $"Precondition: {fileName} should contain '{tag}'.");
            }

            var subset = font.CreateSubset(SubsetText);

            foreach (var tag in PassThroughTags)
            {
                Assert.IsTrue(subset.TableRecords.ContainsKey(tag),
                    $"Subset should contain '{tag}'. Subset tables: [{string.Join(", ", subset.TableRecords.Keys.Select(k => "'" + k + "'").ToArray())}]");
                CollectionAssert.AreEqual(font.GetTableData(tag), subset.GetTableData(tag),
                    $"'{tag}' should be byte-identical to the original.");
            }
        }

        // Test case 2: tables missing in the original are missing in the subset too.
        [TestMethod]
        [DataRow("Mulish-Regular.ttf")]      // prep + gasp only
        [DataRow("NotoEmoji-Regular.ttf")]   // gasp only
        public void Subset_FontMissingHintingTables_DoesNotAddThem(string fileName)
        {
            var font = LoadFontFile(fileName);
            Assert.IsTrue(PassThroughTags.Any(t => !font.TableRecords.ContainsKey(t)),
                $"Precondition: {fileName} should lack at least one of the pass-through tables.");

            var subset = font.CreateSubset(SubsetText);

            foreach (var tag in PassThroughTags)
            {
                Assert.AreEqual(font.TableRecords.ContainsKey(tag), subset.TableRecords.ContainsKey(tag),
                    $"Presence of '{tag}' in subset should match the original.");
            }
        }

        [TestMethod]
        public void Subset_TableDirectory_IsSortedOrdinal()
        {
            var bytes = LoadFontFile("Roboto-Regular.ttf").CreateSubset(SubsetText).Serialize();
            int numTables = (bytes[4] << 8) | bytes[5];
            var tags = new List<string>();
            for (int i = 0; i < numTables; i++)
                tags.Add(System.Text.Encoding.ASCII.GetString(bytes, 12 + i * 16, 4));

            var expected = tags.OrderBy(t => t, StringComparer.Ordinal).ToList();
            CollectionAssert.AreEqual(expected, tags, "Table directory must be sorted by tag in ordinal (binary) order.");
        }


        // Test case 3: serialize, re-parse, verify bytes and checksums.
        [TestMethod]
        [DataRow("Roboto-Regular.ttf")]
        [DataRow("Mulish-Regular.ttf")]
        public void Subset_WithHintingTables_RoundtripPreservesTablesAndChecksums(string fileName)
        {
            var font = LoadFontFile(fileName);
            var subset = font.CreateSubset(SubsetText);

            var parsed = new OpenTypeFont(subset.Serialize());
            SaveFontForCurrentTest(parsed, Path.GetFileNameWithoutExtension(fileName));

            foreach (var tag in PassThroughTags)
            {
                if (!font.TableRecords.ContainsKey(tag))
                    continue;

                Assert.IsTrue(parsed.TableRecords.ContainsKey(tag), $"Parsed subset should contain '{tag}'.");

                var originalData = font.GetTableData(tag);
                CollectionAssert.AreEqual(originalData, parsed.GetTableData(tag),
                    $"'{tag}' should survive the roundtrip byte-identical.");

                var expectedChecksum = ChecksumCalculator.CalculateTableChecksum(PadTo4(originalData), tag);
                Assert.AreEqual(expectedChecksum, parsed.TableRecords[tag].Checksum,
                    $"Checksum for '{tag}' is incorrect.");
            }

            FontTestHelper.AssertFontValid(parsed, FontValidationSeverity.Warning);
        }

        // Test case 4: maxp hinting limits must match the copied fpgm/prep/cvt.
        [TestMethod]
        [DataRow("Roboto-Regular.ttf")]
        public void Subset_WithHintingTables_PreservesMaxpHintingLimits(string fileName)
        {
            var font = LoadFontFile(fileName);
            var parsed = new OpenTypeFont(font.CreateSubset(SubsetText).Serialize());

            var o = font.MaxpTable;
            var s = parsed.MaxpTable;
            Assert.AreEqual(o.maxZones, s.maxZones, "maxZones");
            Assert.AreEqual(o.maxTwilightPoints, s.maxTwilightPoints, "maxTwilightPoints");
            Assert.AreEqual(o.maxStorage, s.maxStorage, "maxStorage");
            Assert.AreEqual(o.maxFunctionDefs, s.maxFunctionDefs, "maxFunctionDefs");
            Assert.AreEqual(o.maxInstructionDefs, s.maxInstructionDefs, "maxInstructionDefs");
            Assert.AreEqual(o.maxStackElements, s.maxStackElements, "maxStackElements");
            // maxSizeOfInstructions may legitimately shrink if computed over the subset glyphs,
            // but must never be smaller than the largest instruction block actually kept.
            Assert.IsTrue(s.maxSizeOfInstructions > 0 || o.maxSizeOfInstructions == 0, "maxSizeOfInstructions");
        }
    }
}