/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/22/2025         EPPlus Software AB           TTF reading tests
 *************************************************************************************************/
using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Scanner;
using EPPlus.Fonts.OpenType.Tests.Helpers;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System.Diagnostics;

namespace EPPlus.Fonts.OpenType.Tests.Reading
{
    [TestClass]
    public sealed class TtfReadingTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void ReadRobotoRegularTtf()
        {
            OpenTypeFont? font = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular);
            Assert.IsNotNull(font);
            var cmap = font.CmapTable;
            Assert.AreEqual("Roboto", font.FullName);
            Assert.AreEqual("Regular", font.SubFamily);
            Assert.AreEqual(1295, font.GlyfTable.Glyphs.Count);
        }

        [TestMethod]
        public void ReadRobotoBoldTtf()
        {
            Stopwatch sw = Stopwatch.StartNew();
            sw.Start();
            OpenTypeFont? font = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Bold);
            sw.Stop();
            var ms = sw.ElapsedMilliseconds;
            Assert.IsNotNull(font);
            Assert.AreEqual("Roboto Bold", font.FullName);
            Assert.AreEqual("Bold", font.SubFamily);
        }

        [TestMethod]
        public void ReadSourceSans3Otf()
        {
            OpenTypeFont? font = TestFolderEngine.LoadFont("Source Sans 3", FontSubFamily.Regular);
            Assert.IsNotNull(font);
            Assert.AreEqual("Source Sans 3", font.FullName);
            Assert.AreEqual("Regular", font.SubFamily);
        }

        struct LicenseDataHolder()
        {
            public string? FontName;
            public ushort LicenseType;
            public string? LTypeString;
        }

        /// <summary>
        /// Indicates font embedding licensing rights for the font. The interpretation of flags is as follows:
        /// 0: Installable embedding: the font may be embedded, and may be permanently installed for use on a remote systems, or for use by other users.
        /// 2: Restricted License embedding: the font must not be modified, embedded or exchanged in any manner without first obtaining explicit permission of the legal owner.
        /// 4: Preview & Print embedding: the font may be embedded, and may be temporarily loaded on other systems for purposes of viewing or printing the document. Documents containing Preview & Print fonts must be opened "read-only"; no edits can be applied to the document.
        /// 8: Editable embedding: the font may be embedded, and may be temporarily loaded on other systems. As with Preview & Print embedding, documents containing Editable fonts may be opened for reading. In addition, editing is permitted, including ability to format new text using the embedded font, and changes may be saved.
        /// </summary>
        string GetFsString(ushort fsId)
        {
            switch (fsId)
            {
                case 0:
                    return "Installable Embedding";
                case 2:
                    return "Restricted Licence Embedding";
                case 4:
                    return "Preview & Print Embedding";
                case 8:
                    return "Editable Embedding";
                default:
                    return $"UNKNOWN VALUE: '{fsId}' POTENTIALLY CORRUPT FONT";
            }
        }

        [TestMethod]
        public void ReadTccGothic()
        {
            OpenTypeFont? font = TestFolderEngine.LoadFont("BIZ UDGothic", FontSubFamily.Bold);

            Assert.IsNotNull(font);
            Assert.AreEqual("BIZ UDGothic Bold", font.FullName);
            Assert.AreEqual("Bold", font.SubFamily);
        }

        [TestMethod]
        public void ReadGsubTable()
        {
            var font = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular);
            var gSub = font.GsubTable;
            Assert.IsNotNull(gSub);
        }

        [TestMethod]
        public void ReadSixFonts()
        {
            RequireFont(SystemFontsEngine, "Aptos Narrow");
            RequireFont(SystemFontsEngine, "Times New Roman");
            RequireFont(SystemFontsEngine, "Calibri", FontSubFamily.Italic);

            OpenTypeFont ? gothic = TestFolderEngine.LoadFont("BIZ UDGothic", FontSubFamily.Bold);
            OpenTypeFont? calibri = SystemFontsEngine.LoadFont("Calibri", FontSubFamily.Italic);
            OpenTypeFont? aptos = SystemFontsEngine.LoadFont("Aptos Narrow", FontSubFamily.Bold);
            OpenTypeFont? timesNewRoman = SystemFontsEngine.LoadFont("Times New Roman", FontSubFamily.Regular);
            OpenTypeFont? SS3 = TestFolderEngine.LoadFont("Source Sans 3", FontSubFamily.Bold);

            Assert.IsNotNull(gothic);
            Assert.AreEqual("BIZ UDGothic Bold", gothic.FullName);
            Assert.AreEqual("Bold", gothic.SubFamily);

            Assert.IsNotNull(calibri);
            Assert.AreEqual("Calibri Italic", calibri.FullName);
            Assert.AreEqual("Italic", calibri.SubFamily);

            Assert.IsNotNull(aptos);
            Assert.AreEqual("Aptos Narrow Bold", aptos.FullName);
            Assert.AreEqual("Bold", aptos.SubFamily);

            Assert.IsNotNull(timesNewRoman);
            Assert.AreEqual("Times New Roman", timesNewRoman.FullName);
            Assert.AreEqual("Regular", timesNewRoman.SubFamily);

            Assert.IsNotNull(SS3);
            Assert.AreEqual("Source Sans 3", SS3.FullName);
            Assert.AreEqual("Regular", SS3.SubFamily);
        }

        [TestMethod]
        public void ReadAllOTFFonts()
        {
            List<OpenTypeFont> allFontsList = OpenTypeFonts.GetAllBaseFontData(FontFolders, true, Scanner.FontFormat.Otf);

            List<LicenseDataHolder> dataHolder = new List<LicenseDataHolder>();

            for (int i = 0; i < allFontsList.Count; i++)
            {
                LicenseDataHolder dataHolderItem = new LicenseDataHolder()
                {
                    FontName = allFontsList[i].FullName,
                    LicenseType = allFontsList[i].Os2Table.fsType,
                    LTypeString = GetFsString(allFontsList[i].Os2Table.fsType)
                };

                dataHolder.Add(dataHolderItem);
                Assert.AreEqual(Scanner.FontFormat.Otf, allFontsList[i].Format);
            }

            var fontsThatCannotBeEmbedded = dataHolder.Where(x => x.LicenseType == 2);

            Assert.AreEqual(0, fontsThatCannotBeEmbedded.Count());
        }

        [TestMethod]
        public void ReadAllTestTTFFontsAndVerifyUseIsOkay()
        {
            List<OpenTypeFont> allFontsList = OpenTypeFonts.GetAllBaseFontData(FontFolders, false, Scanner.FontFormat.Ttf);

            List<LicenseDataHolder> dataHolder = new List<LicenseDataHolder>();

            for (int i = 0; i < allFontsList.Count; i++)
            {
                LicenseDataHolder dataHolderItem = new LicenseDataHolder()
                {
                    FontName = allFontsList[i].FullName,
                    LicenseType = allFontsList[i].Os2Table.fsType,
                    LTypeString = GetFsString(allFontsList[i].Os2Table.fsType)
                };

                dataHolder.Add(dataHolderItem);
                Assert.AreEqual(Scanner.FontFormat.Ttf, allFontsList[i].Format);
            }

            //Ensure all fonts in the font folder are free to install/use
            var fontsThatAreInstallableEmbedded = dataHolder.Where(x => x.LicenseType == 0);

            Assert.AreEqual(fontsThatAreInstallableEmbedded.Count(), dataHolder.Count());
        }

        [TestMethod, Ignore("This test takes a long time and should not run in quick regression tests")]
        public void ReadAllTTFFonts()
        {
            List<OpenTypeFont> allFontsList = OpenTypeFonts.GetAllBaseFontData(FontFolders, true, Scanner.FontFormat.Ttf);

            List<LicenseDataHolder> dataHolder = new List<LicenseDataHolder>();

            for (int i = 0; i < allFontsList.Count; i++)
            {
                LicenseDataHolder dataHolderItem = new LicenseDataHolder()
                {
                    FontName = allFontsList[i].FullName,
                    LicenseType = allFontsList[i].Os2Table.fsType,
                    LTypeString = GetFsString(allFontsList[i].Os2Table.fsType)
                };

                dataHolder.Add(dataHolderItem);
                Assert.AreEqual(Scanner.FontFormat.Ttf, allFontsList[i].Format);
            }

            var fontsThatCannotBeEmbedded = dataHolder.Where(x => x.LicenseType == 2);

            Assert.AreEqual(0, fontsThatCannotBeEmbedded.Count());
        }

        [TestMethod, Ignore("This test takes a long time and should not run in quick regression tests")]
        public void ReadAllFonts()
        {
            var sw = new Stopwatch();
            sw.Start();
            List<OpenTypeFont> allFontsList = OpenTypeFonts.GetAllBaseFontData(FontFolders, true);
            sw.Stop();

            Trace.WriteLine(sw.ElapsedMilliseconds);
        }

        [TestMethod]
        public void TestWrapText()
        {
            string fontName = "Aptos Narrow";
            string testStr = "hello the most";
            double fontSize = 11.0d;
            double MaxPixelWidth = 52d;

            MeasurementFont mf = new MeasurementFont()
            {
                FontFamily = fontName,
                Size = (float)fontSize,
                Style = MeasurementFontStyles.Regular
            };

            var fontMeasurer = OpenTypeFonts.GetTextLayoutEngineForFont(mf);
            var strings = fontMeasurer.WrapText(testStr, mf.Size, MaxPixelWidth);

            Assert.AreEqual("hello the", strings[0]);
            Assert.AreEqual("most", strings[1]);
        }

        [TestMethod]
        public void TestWrapTextWhenLineBreaks()
        {
            string fontName = "Aptos Narrow";
            string testStr = "hello\r\n the\r\n most";
            double fontSize = 11.0d;
            double MaxPixelWidth = 52d;

            RequireFont(SystemFontsEngine, fontName);

            MeasurementFont mf = new MeasurementFont()
            {
                FontFamily = fontName,
                Size = (float)fontSize,
                Style = MeasurementFontStyles.Regular
            };

            var fontMeasurer = SystemFontsEngine.GetTextLayoutEngineForFont(mf);
            var strings = fontMeasurer.WrapText(testStr, mf.Size, MaxPixelWidth);

            Assert.AreEqual("hello", strings[0]);
            Assert.AreEqual(" the", strings[1]);
            Assert.AreEqual(" most", strings[2]);
        }
    }
}