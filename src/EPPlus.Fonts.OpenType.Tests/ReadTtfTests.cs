using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.Scanner;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System.Diagnostics;
using static System.Net.Mime.MediaTypeNames;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public sealed class ReadTtfTests
    {
        private static string _fontFolder = string.Empty;
        private static List<string> _fontFolders = new List<string>();

        [ClassInitialize]
        public static void Initialize(TestContext testContext)
        {
            _fontFolder = Path.Combine(AppContext.BaseDirectory, "Fonts");
            _fontFolders.Clear();
            _fontFolders.Add(_fontFolder);
        }

        [TestMethod]
        public void ReadRobotoRegularTtf()
        {
            TtfFont? font = OpenTypeFonts.GetFontData(_fontFolders, "Roboto", "Regular", false);
            Assert.IsNotNull(font);
            var cmap = font.CmapTable;
            Assert.AreEqual("Roboto", font.FullName);
            Assert.AreEqual("Regular", font.SubFamily);
            Assert.AreEqual(1295, font.GlyphTable.Glyphs.Length);
        }

        [TestMethod]
        public void ReadSourceSans3Otf()
        {
            OpenTypeFont? font = OpenTypeFonts.GetFontDataOpen(_fontFolders, "Source Sans 3", "Regular", false);
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
        /// 4: Preview & Print embedding: the font may be embedded, and may be temporarily loaded on other systems for purposes of viewing or printing the document. Documents containing Preview & Print fonts must be opened “read-only”; no edits can be applied to the document.
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
            OpenTypeFont? font = OpenTypeFonts.GetFontDataOpen(_fontFolders, "BIZ UDGothic", "Bold", true);

            Assert.IsNotNull(font);
            Assert.AreEqual("BIZ UDGothic Bold", font.FullName);
            Assert.AreEqual("Bold", font.SubFamily);
        }

        [TestMethod]
        public void ReadSixFonts()
        {
            List<string> test = new List<string> { Path.Combine(Directory.GetCurrentDirectory(), "Fonts") };
            OpenTypeFont? gothic = OpenTypeFonts.GetFontDataOpen(test, "BIZ UDGothic", "Bold", true);
            OpenTypeFont? calibri = OpenTypeFonts.GetFontDataOpen(test, "Calibri", "Italic", true);
            OpenTypeFont? aptos = OpenTypeFonts.GetFontDataOpen(test, "Aptos Narrow", "Bold", true);
            OpenTypeFont? timesNewRoman = OpenTypeFonts.GetFontDataOpen(test, "Times New Roman", "Regular", true);
            OpenTypeFont? SS3 = OpenTypeFonts.GetFontDataOpen(_fontFolders, "Source Sans 3", "Bold", false);

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
            List<OpenTypeFont> allFontsList = OpenTypeFonts.GetAllBaseFontData(_fontFolders, true, Scanner.FontFormat.Otf);

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
        public void ReadAllTTFFonts()
        {
            List<OpenTypeFont> allFontsList = OpenTypeFonts.GetAllBaseFontData(_fontFolders, true, Scanner.FontFormat.Ttf);

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

        [TestMethod]
        public void ReadAllFonts()
        {
            var sw = new Stopwatch();
            sw.Start();
            List<OpenTypeFont> allFontsList = OpenTypeFonts.GetAllBaseFontData(_fontFolders, true);
            sw.Stop();

            Trace.WriteLine(sw.ElapsedMilliseconds);


            ////List<LicenseDataHolder> dataHolder = new List<LicenseDataHolder>();

            ////for (int i = 0; i < allFontsList.Count; i++)
            ////{
            ////    LicenseDataHolder dataHolderItem = new LicenseDataHolder()
            ////    {
            ////        FontName = allFontsList[i].FullName,
            ////        LicenseType = allFontsList[i].Os2Table.fsType,
            ////        LTypeString = GetFsString(allFontsList[i].Os2Table.fsType)
            ////    };

            ////    dataHolder.Add(dataHolderItem);
            ////}

            ////var fontsThatCannotBeEmbedded = dataHolder.Where(x => x.LicenseType == 2);

            //Assert.AreEqual(0, fontsThatCannotBeEmbedded.Count());
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


            FontMeasurerTrueType fontMeasurer = new FontMeasurerTrueType(fontSize, fontName);
            var strings = fontMeasurer.MeasureAndWrapText(testStr, mf, MaxPixelWidth);

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

            MeasurementFont mf = new MeasurementFont()
            {
                FontFamily = fontName,
                Size = (float)fontSize,
                Style = MeasurementFontStyles.Regular
            };


            FontMeasurerTrueType fontMeasurer = new FontMeasurerTrueType(fontSize, fontName);
            var strings = fontMeasurer.MeasureAndWrapText(testStr, mf, MaxPixelWidth);

            Assert.AreEqual("hello", strings[0]);
            Assert.AreEqual(" the", strings[1]);
            Assert.AreEqual(" most", strings[2]);
        }

    }
}
