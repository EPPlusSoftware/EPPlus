/*******************************************************************************
 * You may amend and distribute as you like, but don't remove this header!
 *
 * Required Notice: Copyright (C) EPPlus Software AB. 
 * https://epplussoftware.com
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.

 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  
 * See the GNU Lesser General Public License for more details.
 *
 * The GNU Lesser General Public License can be viewed at http://www.opensource.org/licenses/lgpl-license.php
 * If you unfamiliar with this license or have questions about it, here is an http://www.gnu.org/licenses/gpl-faq.html
 *
 * All code and executables are provided "" as is "" with no warranty either express or implied. 
 * The author accepts no liability for any damage or loss of business that this product may cause.
 *
 * Code change notes:
 * 
  Date               Author                       Change
 *******************************************************************************
  08/25/2026         EPPlus Software AB       Regression tests: header/footer sections
                                                must parse regardless of their order
 *******************************************************************************/
using System;
using System.IO;
using System.Xml;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;

namespace EPPlusTest.Core.Worksheet
{
    /// <summary>
    /// Regression tests for header/footer sections being lost when they are not stored in
    /// Left, Center, Right order.
    ///
    /// The ExcelHeaderFooterText constructor takes the first section code from the first two
    /// characters of the raw string, then scans for further section codes. Before the fix its
    /// scan only recognized "&amp;C" and "&amp;R" - never "&amp;L" - so a "&amp;L" appearing
    /// anywhere other than at position 0 was not treated as the start of a new section. Its
    /// content was swallowed into the preceding section, and then discarded when that section
    /// was normalized by ReadHeaderFooterFormat/WriteHeaderFooterFormat, taking any "&amp;G"
    /// picture placeholder with it. The picture survived in the VML collection, so
    /// HeaderFooter.Pictures.Count was unchanged, but nothing referenced it any more and it
    /// stopped rendering in Excel.
    ///
    /// Excel writes the sections in the order the user created them, which is frequently not
    /// Left, Center, Right - so this affected ordinary files produced by Excel itself.
    ///
    /// Reported case (oddfooter-left-corruption-template.xlsx), oddFooter raw value:
    ///     &amp;C&amp;"-,Bold"&amp;12{FORMID}&amp;"-,Regular"&amp;11\n&amp;L&amp;G&amp;R&amp;G
    /// which before the fix was persisted as:
    ///     &amp;C&amp;"-,Bold"&amp;12{FORMID}&amp;"-,Regular"&amp;11\n&amp;R&amp;G
    ///
    /// The same file's oddHeader ("&amp;L&amp;G&amp;R&amp;G") was unaffected, both because it
    /// starts with "&amp;L" and because Save() only rewrites a node whose backing field is
    /// non-null - and the reported repro only ever touched OddFooter.
    /// </summary>
    [TestClass]
    public class HeaderFooterSectionOrderTests : TestBase
    {
        /// <summary>The customer's exact oddFooter value. Section order: Center, Left, Right.</summary>
        private const string CustomerOddFooter = "&C&\"-,Bold\"&12{FORMID}&\"-,Regular\"&11\n&L&G&R&G";

        /// <summary>The customer's exact oddHeader value. Section order: Left, Right.</summary>
        private const string CustomerOddHeader = "&L&G&R&G";

        [ClassInitialize]
        public static void Init(TestContext testContext)
        {
            InitBase();
        }

        [TestMethod]
        public void CustomerOddFooter_SurvivesRoundTrip_WhenOddFooterIsTouched()
        {
            // The exact scenario from the support ticket: a footer whose sections are stored
            // in Center, Left, Right order, with a single read of a header/footer property as
            // the only interaction before saving.
            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                SetRawHeaderFooterNode(ws, "oddFooter", CustomerOddFooter);
                SetRawHeaderFooterNode(ws, "oddHeader", CustomerOddHeader);

                // The single interaction from the customer's repro script.
                var _ = ws.HeaderFooter.OddFooter.LeftAlignedText;

                var persistedFooter = SaveAndReadRawNode(pkg, "oddFooter");
                var persistedHeader = SaveAndReadRawNode(pkg, "oddHeader");

                StringAssert.Contains(persistedFooter, "&L" + ExcelHeaderFooter.Image,
                    "The Left section's picture placeholder must survive the round trip. " +
                    $"Persisted footer was: \"{Escape(persistedFooter)}\".");

                Assert.AreEqual(CustomerOddHeader, persistedHeader,
                    "The untouched oddHeader must be unchanged by the save.");
            }
        }

        [TestMethod]
        public void LeftSection_IsParsed_RegardlessOfSectionOrder()
        {
            // Every one of these holds the same three logical sections, only reordered. The
            // Left section always contains a picture placeholder, and must always come back.
            var variants = new[]
            {
                new { Name = "L,C,R", Raw = "&L&G&CCenterText&R&G" },
                new { Name = "C,L,R", Raw = "&CCenterText&L&G&R&G" },
                new { Name = "R,L,C", Raw = "&R&G&L&G&CCenterText" },
                new { Name = "C,R,L", Raw = "&CCenterText&R&G&L&G" },
                new { Name = "R,C,L", Raw = "&R&G&CCenterText&L&G" },
                new { Name = "L,R",   Raw = "&L&G&R&G" },
                new { Name = "C,L",   Raw = "&CCenterText&L&G" },
                new { Name = "R,L",   Raw = "&R&G&L&G" },
            };

            foreach (var v in variants)
            {
                using (var pkg = new ExcelPackage())
                {
                    var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                    SetRawHeaderFooterNode(ws, "oddFooter", v.Raw);

                    var oddFooter = ws.HeaderFooter.OddFooter;

                    StringAssert.Contains(oddFooter.LeftAlignedText, ExcelHeaderFooter.Image,
                        $"[{v.Name}] The Left section's '&G' placeholder was not parsed from " +
                        $"\"{Escape(v.Raw)}\". Left parsed as \"{Escape(oddFooter.LeftAlignedText)}\".");

                    // Guard against the content merely being relocated into another section.
                    Assert.IsFalse(Contains(oddFooter.CenteredText, "&L"),
                        $"[{v.Name}] The Center section swallowed a stray '&L': " +
                        $"\"{Escape(oddFooter.CenteredText)}\".");
                    Assert.IsFalse(Contains(oddFooter.RightAlignedText, "&L"),
                        $"[{v.Name}] The Right section swallowed a stray '&L': " +
                        $"\"{Escape(oddFooter.RightAlignedText)}\".");
                }
            }
        }

        [TestMethod]
        public void AllSections_SurviveRoundTrip_RegardlessOfSectionOrder()
        {
            // As above, but verifying what actually reaches the file. Section order is not
            // required to be preserved - only the content of each section.
            var variants = new[]
            {
                new { Name = "L,C,R", Raw = "&L&G&CCenterText&R&G" },
                new { Name = "C,L,R", Raw = "&CCenterText&L&G&R&G" },
                new { Name = "R,L,C", Raw = "&R&G&L&G&CCenterText" },
                new { Name = "C,R,L", Raw = "&CCenterText&R&G&L&G" },
                new { Name = "R,C,L", Raw = "&R&G&CCenterText&L&G" },
            };

            foreach (var v in variants)
            {
                using (var pkg = new ExcelPackage())
                {
                    var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                    SetRawHeaderFooterNode(ws, "oddFooter", v.Raw);

                    // Touch the object so Save() rewrites the node.
                    var _ = ws.HeaderFooter.OddFooter.CenteredText;

                    var persisted = SaveAndReadRawNode(pkg, "oddFooter");

                    StringAssert.Contains(persisted, "&L" + ExcelHeaderFooter.Image,
                        $"[{v.Name}] Left section lost. Raw was \"{Escape(v.Raw)}\", " +
                        $"persisted \"{Escape(persisted)}\".");
                    StringAssert.Contains(persisted, "&CCenterText",
                        $"[{v.Name}] Center section lost. Persisted \"{Escape(persisted)}\".");
                    StringAssert.Contains(persisted, "&R" + ExcelHeaderFooter.Image,
                        $"[{v.Name}] Right section lost. Persisted \"{Escape(persisted)}\".");
                }
            }
        }

        [TestMethod]
        public void EmptySectionBetweenTwoSections_DoesNotConsumeFollowingSectionCode()
        {
            // Covers the "pos = startPos - 1" part of the fix. With the previous
            // "pos = startPos" the loop's pos++ skipped the character at startPos, so a
            // section code starting immediately after a consumed code - i.e. an empty
            // section - was missed, and the following section was swallowed by it.
            //
            // The empty section must sit in the MIDDLE for this to bite: the first section
            // code is taken outside the loop, so no skip has happened yet at that point.
            // "&C&L&G&R&G" therefore parses correctly even without the fix and would not
            // catch a regression here - the empty Center has to follow a match made inside
            // the loop, as below.
            const string raw = "&L&G&C&R&G";

            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                SetRawHeaderFooterNode(ws, "oddFooter", raw);

                var oddFooter = ws.HeaderFooter.OddFooter;

                Console.WriteLine($"Raw    : \"{Escape(raw)}\"");
                Console.WriteLine($"Left   : \"{Escape(oddFooter.LeftAlignedText)}\"");
                Console.WriteLine($"Center : \"{Escape(oddFooter.CenteredText)}\"");
                Console.WriteLine($"Right  : \"{Escape(oddFooter.RightAlignedText)}\"");

                StringAssert.Contains(oddFooter.LeftAlignedText, ExcelHeaderFooter.Image,
                    $"Left section not parsed from \"{Escape(raw)}\".");

                // The empty Center must not swallow the Right section's code.
                Assert.IsFalse(Contains(oddFooter.CenteredText, "&R"),
                    $"The empty Center section swallowed the following '&R': " +
                    $"\"{Escape(oddFooter.CenteredText)}\".");

                // This is the assertion that should fail if "pos = startPos - 1" is reverted:
                // Right is never set, so RightAlignedText comes back as just "&R".
                StringAssert.Contains(oddFooter.RightAlignedText, ExcelHeaderFooter.Image,
                    $"Right section not parsed from \"{Escape(raw)}\" - it was most likely " +
                    $"consumed by the empty Center section. Right parsed as " +
                    $"\"{Escape(oddFooter.RightAlignedText)}\".");
            }
        }

        [TestMethod]
        public void SectionCodeAtEndOfString_IsRecognized()
        {
            // Covers the "text.Length - 1" part of the fix. The previous "text.Length - 2"
            // bound meant a section code occupying the final two characters was never seen,
            // so it was swallowed into the preceding section instead of starting an empty one.
            const string raw = "&L&G&R";

            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                SetRawHeaderFooterNode(ws, "oddFooter", raw);

                var oddFooter = ws.HeaderFooter.OddFooter;

                Assert.IsFalse(Contains(oddFooter.LeftAlignedText, "&R"),
                    $"The trailing '&R' was swallowed into the Left section: " +
                    $"\"{Escape(oddFooter.LeftAlignedText)}\".");
                StringAssert.Contains(oddFooter.LeftAlignedText, ExcelHeaderFooter.Image,
                    "The Left section's picture placeholder should still be parsed.");
            }
        }

        [TestMethod]
        public void HeaderSections_AreParsed_RegardlessOfSectionOrder()
        {
            // The same parsing path backs OddHeader, so it needs the same coverage - the
            // reported case simply never touched the header.
            const string raw = "&CHeaderCenter&L&G&R&G";

            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                SetRawHeaderFooterNode(ws, "oddHeader", raw);

                var oddHeader = ws.HeaderFooter.OddHeader;

                StringAssert.Contains(oddHeader.LeftAlignedText, ExcelHeaderFooter.Image,
                    $"Left header section not parsed from \"{Escape(raw)}\". " +
                    $"Left parsed as \"{Escape(oddHeader.LeftAlignedText)}\".");
            }
        }

        [TestMethod]
        public void PictureCount_IsUnchanged_ByHeaderFooterTextRoundTrip()
        {
            // The reported symptom that made this hard to spot: the picture object itself was
            // never lost, only the text reference to it, so Pictures.Count kept reporting the
            // original value. This pins that behavior down so a future change cannot start
            // silently dropping the VML pictures instead.
            using (var pkg = new ExcelPackage())
            {
                var ws = pkg.Workbook.Worksheets.Add("Sheet1");
                SetRawHeaderFooterNode(ws, "oddFooter", CustomerOddFooter);

                var countBefore = ws.HeaderFooter.Pictures.Count;
                var _ = ws.HeaderFooter.OddFooter.LeftAlignedText;

                using (var stream = new MemoryStream())
                {
                    pkg.SaveAs(stream);
                    using (var reloaded = new ExcelPackage(stream))
                    {
                        Assert.AreEqual(countBefore,
                            reloaded.Workbook.Worksheets[0].HeaderFooter.Pictures.Count,
                            "HeaderFooter.Pictures.Count changed across the round trip.");
                    }
                }
            }
        }

        #region Helpers

        private static string Escape(string s)
        {
            return s == null ? "<null>" : s.Replace("\n", "\\n").Replace("\r", "\\r");
        }

        private static bool Contains(string haystack, string needle)
        {
            return haystack != null && haystack.Contains(needle);
        }

        private static XmlNamespaceManager GetNsm(ExcelWorksheet ws)
        {
            var nsm = new XmlNamespaceManager(ws.WorksheetXml.NameTable);
            nsm.AddNamespace("d", ExcelPackage.schemaMain);
            return nsm;
        }

        /// <summary>
        /// Writes a raw string straight into d:headerFooter/d:{nodeName}, so section ordering
        /// is under the test's control instead of EPPlus's own always-Left-Center-Right
        /// authoring order - which is what makes these orderings reachable at all.
        /// </summary>
        private static void SetRawHeaderFooterNode(ExcelWorksheet ws, string nodeName, string rawText)
        {
            var nsm = GetNsm(ws);
            var wsNode = ws.WorksheetXml.SelectSingleNode("d:worksheet", nsm);

            var hfNode = wsNode.SelectSingleNode("d:headerFooter", nsm);
            if (hfNode == null)
            {
                hfNode = ws.WorksheetXml.CreateElement("headerFooter", ExcelPackage.schemaMain);
                wsNode.AppendChild(hfNode);
            }

            var node = hfNode.SelectSingleNode("d:" + nodeName, nsm);
            if (node == null)
            {
                node = ws.WorksheetXml.CreateElement(nodeName, ExcelPackage.schemaMain);
                hfNode.AppendChild(node);
            }
            node.InnerText = rawText;
        }

        /// <summary>
        /// Saves the package to a stream, reloads it, and returns the raw text of the
        /// requested header/footer node as persisted - i.e. what Excel would read.
        /// </summary>
        private static string SaveAndReadRawNode(ExcelPackage pkg, string nodeName)
        {
            using (var stream = new MemoryStream())
            {
                pkg.SaveAs(stream);

                using (var reloaded = new ExcelPackage(stream))
                {
                    var ws = reloaded.Workbook.Worksheets[0];
                    var nsm = GetNsm(ws);
                    var node = ws.WorksheetXml.SelectSingleNode(
                        "d:worksheet/d:headerFooter/d:" + nodeName, nsm);
                    return node == null ? null : node.InnerText;
                }
            }
        }

        #endregion
    }
}