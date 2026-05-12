using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.TextShaping;
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class TextFragmentCollectionTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        protected override void ConfigureResolver(bool searchSystemDirectories = true)
        {
            base.ConfigureResolver(searchSystemDirectories);
        }

        // ============================================================================
        // DIAGNOSTIC INSTRUMENTATION — paste this entire test method into your test
        // class, replacing the existing EnsureTextFragmentsAndWrapperWorkCorrectlyForLongParagraphs.
        //
        // It does what the original test does, plus prints diagnostic snapshots before
        // and after the shaper is created. The output goes to Debug.WriteLine — visible
        // in the Test Explorer "Output" panel for the test, or in any attached debugger.
        //
        // HOW TO USE:
        //   1. Paste this into your test file (replace the existing method).
        //   2. Run the full test suite. The test will pass first time, fail second time.
        //   3. After running the suite TWICE, look at the Output for the test in run 2.
        //      Compare the diagnostic output between run 1 and run 2.
        //   4. Copy/paste both outputs back to chat for analysis.
        // ============================================================================



        [TestMethod]
        public void EnsureTextFragmentsAndWrapperWorkCorrectlyForLongParagraphs()
        {
            DumpResolverState("BEFORE shaper creation");

            var shaper = OpenTypeFonts.GetTextShaper("Aptos Narrow", FontSubFamily.Regular);

            DumpResolverState("AFTER shaper creation");
            DumpShaperState("shaper", shaper);

            // ----- DIAGNOSTIC: inspect the OpenTypeFont via OpenTypeFonts.LoadFont -----
            // If LoadFont returns a cached instance, two runs see the same object —
            // and any mutation of that object's tables will show up here.
            DumpFontState("LoadFont Aptos Narrow Regular");

            var layout = new TextLayoutEngine(shaper);

            var outputLines = layout.WrapText(
                "Hello World! a b c d e f g h i j k l m n o p q r s t u v w x y z \r\n",
                28f,
                225);

            System.Console.WriteLine("[DIAG] === Output lines: " + outputLines.Count + " ===");
            for (int i = 0; i < outputLines.Count; i++)
            {
                System.Console.WriteLine(
                    "[DIAG] line[" + i + "] (len=" + outputLines[i].Length + ") = \"" + outputLines[i] + "\"");
            }
            System.Console.WriteLine("[DIAG] === END ===");

            Assert.AreEqual("Hello World! a b c d", outputLines[0]);
            Assert.AreEqual("e f g h i j k l m n o p q", outputLines[1]);
            Assert.AreEqual("r s t u v w x y z ", outputLines[2]);
        }

        // ============================================================================
        // Helpers — paste these into the same test class.
        // ============================================================================

        private static void DumpResolverState(string label)
        {
            System.Console.WriteLine("[DIAG] ----- " + label + " -----");

            var resolver = new EPPlus.Fonts.OpenType.FontResolver.DefaultFontResolver();
            var bytes = resolver.ResolveFont("Aptos Narrow", OfficeOpenXml.Interfaces.Fonts.FontSubFamily.Regular);

            System.Console.WriteLine("[DIAG]   resolved bytes: length=" + bytes.Length
                + " sha1=" + Sha1Short(bytes)
                + " head=" + HexHead(bytes, 16));

            try
            {
                var font = EPPlus.Fonts.OpenType.OpenTypeFonts.GetFromBytes(bytes);
                System.Console.WriteLine("[DIAG]   parsed family=" + font.NameTable.GetFamilyName()
                    + " subfamily=" + font.NameTable.GetSubfamilyEnum());
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine("[DIAG]   parse FAILED: " + ex.GetType().Name + ": " + ex.Message);
            }

            var scanner = new DefaultFontScanner();
            var face = scanner.FindBestMatch(
                new System.Collections.Generic.List<string>(),
                "Aptos Narrow",
                OfficeOpenXml.Interfaces.Fonts.FontSubFamily.Regular,
                true);

            if (face == null)
                System.Console.WriteLine("[DIAG]   scanner.FindBestMatch returned null");
            else
                System.Console.WriteLine("[DIAG]   scanner.FindBestMatch:"
                    + " family=" + face.FamilyName
                    + " subfamily=" + face.Subfamily
                    + " path=" + System.IO.Path.GetFileName(face.FilePath ?? "(null)")
                    + " IsExactMatch=" + face.IsExactMatch);
        }

        private static void DumpFontState(string label)
        {
            System.Console.WriteLine("[DIAG] ----- " + label + " -----");

            try
            {
                var font = EPPlus.Fonts.OpenType.OpenTypeFonts.LoadFont(
                    "Aptos Narrow",
                    OfficeOpenXml.Interfaces.Fonts.FontSubFamily.Regular);

                // Identity — same instance across calls?
                System.Console.WriteLine("[DIAG]   font instance hash=" + System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(font));
                System.Console.WriteLine("[DIAG]   font.RawData length=" + (font.RawData == null ? -1 : font.RawData.Length));
                if (font.RawData != null && font.RawData.Length > 0)
                    System.Console.WriteLine("[DIAG]   font.RawData head=" + HexHead(font.RawData, 16));

                // Glyph mapping for the characters that matter for the failing test.
                // 'e' is the character that just barely fits / doesn't fit in line[0].
                DumpGlyphAndWidth(font, 'H', "H");
                DumpGlyphAndWidth(font, 'e', "e");
                DumpGlyphAndWidth(font, 'l', "l");
                DumpGlyphAndWidth(font, 'o', "o");
                DumpGlyphAndWidth(font, ' ', "space");
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine("[DIAG]   FAILED: " + ex.GetType().Name + ": " + ex.Message);
            }
        }

        private static void DumpGlyphAndWidth(EPPlus.Fonts.OpenType.OpenTypeFont font, char ch, string label)
        {
            try
            {
                ushort glyphId;
                bool found = font.CmapTable.TryGetGlyphId((uint)ch, out glyphId);
                if (!found)
                {
                    System.Console.WriteLine("[DIAG]   '" + label + "' (U+" + ((int)ch).ToString("X4") + "): no cmap entry");
                    return;
                }

                // Try to read advance width from hmtx. If the property/method on your OpenTypeFont
                // is named differently, adjust this call.
                int advance = -1;
                try
                {
                    advance = font.HmtxTable.GetAdvanceWidth(glyphId);
                }
                catch (System.Exception ex)
                {
                    System.Console.WriteLine("[DIAG]   hmtx lookup failed for glyph " + glyphId + ": " + ex.GetType().Name);
                }

                System.Console.WriteLine("[DIAG]   '" + label + "' (U+" + ((int)ch).ToString("X4") + "): glyphId=" + glyphId + " advanceWidth=" + advance);
            }
            catch (System.Exception ex)
            {
                System.Console.WriteLine("[DIAG]   '" + label + "' FAILED: " + ex.GetType().Name + ": " + ex.Message);
            }
        }

        private static void DumpShaperState(string label, object shaper)
        {
            System.Console.WriteLine("[DIAG] " + label + " type=" + (shaper == null ? "null" : shaper.GetType().Name));
        }

        private static string Sha1Short(byte[] bytes)
        {
            using (var sha = System.Security.Cryptography.SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(bytes);
                var sb = new System.Text.StringBuilder();
                for (int i = 0; i < 8 && i < hash.Length; i++)
                    sb.Append(hash[i].ToString("x2"));
                return sb.ToString();
            }
        }

        private static string HexHead(byte[] bytes, int count)
        {
            var sb = new System.Text.StringBuilder();
            int n = System.Math.Min(count, bytes.Length);
            for (int i = 0; i < n; i++)
            {
                if (i > 0) sb.Append(' ');
                sb.Append(bytes[i].ToString("x2"));
            }
            return sb.ToString();
        }

        //TODO: DOUBLE-CHECK BOLD+ITALIC for narrow later it seems innaccurate
        [TestMethod]
        public void MeasureBold()
        {
            var font2 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };
            string test = "TextBox2";

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();

            var measurer = OpenTypeFonts.GetShaperForFont(font2);
            var widthInPoints = measurer.ShapeLight(test).GetWidthInPoints(font2.Size);

            var inPixels = Math.Round(widthInPoints.PointToPixel(),0,MidpointRounding.AwayFromZero);

            Assert.AreEqual(54, inPixels);
        }

        [TestMethod]
        public void MeasureGoudy()
        {
            var font5 = new MeasurementFont()
            {
                FontFamily = "Goudy Stout",
                Size = 16,
                Style = MeasurementFontStyles.Regular
            };

            var text = "Goudy size";

            var measurer = OpenTypeFonts.GetShaperForFont(font5);

            var widthInPoints = measurer.ShapeLight(text).GetWidthInPoints(16f);

            var inPixels = Math.Round(widthInPoints.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            Assert.AreEqual(237, inPixels);
        }

        [TestMethod]
        public void MeasureWrappedWidthsWithInternalLineBreaks()
        {
            List<string> lstOfRichText = new() { "TextBox\r\na\r\n", "TextBox2", "ra underline", "La Strike", "Goudy size 16", "SvgSize 24" };

            var font1 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Regular
            }; ;

            var font2 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };

            var font3 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Underline
            };

            var font4 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Strikeout
            };

            var font5 = new MeasurementFont()
            {
                FontFamily = "Goudy Stout",
                Size = 16,
                Style = MeasurementFontStyles.Regular
            };


            var font6 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 24,
                Style = MeasurementFontStyles.Regular
            };

            List<MeasurementFont> fonts = new() { font1, font2, font3, font4, font5, font6 };

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            var ttMeasurer = OpenTypeFonts.GetTextLayoutEngineForFont(font1);

            var textFragments = new TextFragmentCollectionSimple(fonts, lstOfRichText);

            var wrappedLines = ttMeasurer.WrapRichTextLines(textFragments, maxSizePoints);

            var line1 = wrappedLines[0];

            var pixels11 = Math.Round(line1.LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeline1 = Math.Round(line1.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            Assert.AreEqual(44, pixels11);
            Assert.AreEqual(pixels11, pixelsWholeline1);

            var line2 = wrappedLines[1];

            var pixels21 = Math.Round(line2.LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeline2 = Math.Round(line2.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            Assert.AreEqual(7, pixels21);
            Assert.AreEqual(pixels21, pixelsWholeline2);

            var line3 = wrappedLines[2];

            var pixels31 = Math.Round(line3.LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels32 = Math.Round(line3.LineFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels33 = Math.Round(line3.LineFragments[2].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine3 = Math.Round(line3.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);

            //~54 px
            Assert.AreEqual(54, pixels31);
            //~70 px aka 51.75pt
            Assert.AreEqual(70, pixels32);
            //~16-17 px This line contains a space at the end
            Assert.AreEqual(17, pixels33);

            //Total Width: ~140
            Assert.AreEqual(140d, pixelsWholeLine3);

            var line4 = wrappedLines[3];

            var pixels41 = Math.Round(line4.LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels42 = Math.Round(line4.LineFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine4 = Math.Round(line4.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            //~34 px
            Assert.AreEqual(33, pixels41);
            // This line contains a space at the end
            Assert.AreEqual(248, pixels42);

            Assert.AreEqual(281, pixelsWholeLine4);

            var line5 = wrappedLines[4];

            var pixels51 = Math.Round(line5.LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels52 = Math.Round(line5.LineFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine5 = Math.Round(line5.Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            Assert.AreEqual(35, pixels51);
            //This line does NOT contain a space at the end
            Assert.AreEqual(134, pixels52);


            Assert.AreEqual(169, pixelsWholeLine5);
        }

        [TestMethod]
        public void MeasureWrappedWidths()
        {
            List<string> lstOfRichText = new() { "TextBox2", "ra underline", "La Strike", "Goudy size 16", "SvgSize 24" };

            var font2 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };

            var font3 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Underline
            };

            var font4 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Strikeout
            };

            var font5 = new MeasurementFont()
            {
                FontFamily = "Goudy Stout",
                Size = 16,
                Style = MeasurementFontStyles.Regular
            };


            var font6 = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 24,
                Style = MeasurementFontStyles.Regular
            };

            List<MeasurementFont> fonts = new() {font2, font3, font4, font5, font6 };

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            var ttMeasurer = OpenTypeFonts.GetTextLayoutEngineForFont(font2);

            var textFragments = new TextFragmentCollectionSimple(fonts, lstOfRichText);

            var wrappedLines = ttMeasurer.WrapRichTextLines(textFragments, maxSizePoints);

            var pixels1 = Math.Round(wrappedLines[0].LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels2 = Math.Round(wrappedLines[0].LineFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels3 = Math.Round(wrappedLines[0].LineFragments[2].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine = Math.Round(wrappedLines[0].Width.PointToPixel(),0, MidpointRounding.AwayFromZero);

            //~54 px
            Assert.AreEqual(54, pixels1);
            //~70 px aka 51.75pt
            Assert.AreEqual(70, pixels2);
            //~16-17 px This line contains a space at the end
            Assert.AreEqual(17, pixels3);

            //Total Width: ~140
            Assert.AreEqual(140d, pixelsWholeLine);

            var pixels21 = Math.Round(wrappedLines[1].LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels22 = Math.Round(wrappedLines[1].LineFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine2 = Math.Round(wrappedLines[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            //~34 px
            Assert.AreEqual(33, pixels21);
            // This line contains a space at the end
            Assert.AreEqual(248, pixels22);

            Assert.AreEqual(281, pixelsWholeLine2);

            var pixels31 = Math.Round(wrappedLines[2].LineFragments[0].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixels32 = Math.Round(wrappedLines[2].LineFragments[1].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            var pixelsWholeLine3 = Math.Round(wrappedLines[2].Width.PointToPixel(), 0, MidpointRounding.AwayFromZero);
            Assert.AreEqual(35, pixels31);
            //This line does NOT contain a space at the end
            Assert.AreEqual(134, pixels32);

            
            Assert.AreEqual(169, pixelsWholeLine3);
        }

        [TestMethod]
        public void CorrectTextLinesAreReturnedWhenSmallMaxWidth()
        {
            var defaultFont = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Regular
            };

            var ttMeasurer = OpenTypeFonts.GetTextLayoutEngineForFont(defaultFont);

            var txt = "This is my text";

            //This should be small enough to put each word on a new row.
            var maxWidth = 21.5d;

            var txtLines = ttMeasurer.WrapText(txt, 11f, maxWidth);

            Assert.AreEqual(4, txtLines.Count);
            Assert.AreEqual("This", txtLines[0]);
            Assert.AreEqual("is", txtLines[1]);
            Assert.AreEqual("my", txtLines[2]);
            Assert.AreEqual("text", txtLines[3]);
        }
    }
}
