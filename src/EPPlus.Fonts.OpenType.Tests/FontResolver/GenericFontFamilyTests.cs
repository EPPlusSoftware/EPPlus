using OfficeOpenXml.Interfaces.Fonts;

namespace EPPlus.Fonts.OpenType.Tests.FontResolver
{
    /// <summary>
    /// Tests for <see cref="OpenTypeFontEngine.GetGenericFamily"/> and the rules in
    /// <see cref="GenericFontFamilyClassifier"/>.
    /// </summary>
    [TestClass]
    public class GenericFontFamilyTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        // -----------------------------------------------------------------------------------------
        // Test fonts
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        public void Roboto_IsSansSerif()
        {
            RequireFont(TestFolderEngine, "Roboto");
            Assert.AreEqual(GenericFontFamily.SansSerif, TestFolderEngine.GetGenericFamily("Roboto"));
        }

        [TestMethod]
        public void BizUdGothic_FixedPitch_IsMonospace()
        {
            RequireFont(TestFolderEngine, "BIZ UDGothic");
            Assert.AreEqual(GenericFontFamily.Monospace, TestFolderEngine.GetGenericFamily("BIZ UDGothic"));
        }

        /// <summary>
        /// EB Garamond has serifs, but sets neither sFamilyClass nor PANOSE. Nothing in the font
        /// says it is a serif font, so it is reported as sans-serif. Documents the limitation.
        /// </summary>
        [TestMethod]
        public void EbGaramond_WithoutClassification_IsSansSerif()
        {
            RequireFont(TestFolderEngine, "EB Garamond");
            Assert.AreEqual(GenericFontFamily.SansSerif, TestFolderEngine.GetGenericFamily("EB Garamond"));
        }

        [TestMethod]
        public void NullOrEmptyName_IsSansSerif()
        {
            Assert.AreEqual(GenericFontFamily.SansSerif, TestFolderEngine.GetGenericFamily(null!));
            Assert.AreEqual(GenericFontFamily.SansSerif, TestFolderEngine.GetGenericFamily(string.Empty));
        }

        // -----------------------------------------------------------------------------------------
        // System fonts (Inconclusive where not installed)
        // -----------------------------------------------------------------------------------------

        [TestMethod]
        [DataRow("Calibri", GenericFontFamily.SansSerif)]
        [DataRow("Arial", GenericFontFamily.SansSerif)]
        [DataRow("Cambria", GenericFontFamily.Serif)]
        [DataRow("Times New Roman", GenericFontFamily.Serif)]
        [DataRow("Georgia", GenericFontFamily.Serif)]
        [DataRow("Consolas", GenericFontFamily.Monospace)]
        [DataRow("Courier New", GenericFontFamily.Monospace)]
        public void SystemFont(string family, GenericFontFamily expected)
        {
            RequireFont(SystemFontsEngine, family);
            Assert.AreEqual(expected, SystemFontsEngine.GetGenericFamily(family));
        }

        // -----------------------------------------------------------------------------------------
        // Classifier rules, on a private copy of Roboto with the relevant fields changed
        // -----------------------------------------------------------------------------------------

        private static OpenTypeFont LoadPrivateRoboto()
        {
            RequireFont(TestFolderEngine, "Roboto");
            // ignoreCache: the fields are changed below and must not leak into the shared engine.
            var font = TestFolderEngine.LoadFont("Roboto", FontSubFamily.Regular, ignoreCache: true);
            font.PostTable.isFixedPitch = 0;
            font.Os2Table.sFamilyClass = 0;
            font.Os2Table.panose = new byte[10];
            return font;
        }

        [TestMethod]
        public void Classifier_NothingSet_IsSansSerif()
        {
            Assert.AreEqual(GenericFontFamily.SansSerif, GenericFontFamilyClassifier.Classify(LoadPrivateRoboto()));
        }

        [TestMethod]
        public void Classifier_IsFixedPitch_IsMonospace_EvenWithSerifClass()
        {
            var font = LoadPrivateRoboto();
            font.PostTable.isFixedPitch = 1;
            font.Os2Table.sFamilyClass = 0x0105;
            Assert.AreEqual(GenericFontFamily.Monospace, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(3)]
        [DataRow(4)]
        [DataRow(5)]
        [DataRow(7)]
        public void Classifier_SerifFamilyClass_IsSerif(int familyClass)
        {
            var font = LoadPrivateRoboto();
            font.Os2Table.sFamilyClass = (short)((familyClass << 8) | 0x03);
            Assert.AreEqual(GenericFontFamily.Serif, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        [DataRow(8)]  // sans serif
        [DataRow(9)]  // ornamentals
        [DataRow(10)] // scripts
        [DataRow(12)] // symbolic
        public void Classifier_OtherFamilyClass_IsSansSerif(int familyClass)
        {
            var font = LoadPrivateRoboto();
            font.Os2Table.sFamilyClass = (short)(familyClass << 8);
            Assert.AreEqual(GenericFontFamily.SansSerif, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        public void Classifier_SerifFamilyClass_WinsOverSansPanose()
        {
            var font = LoadPrivateRoboto();
            font.Os2Table.sFamilyClass = 1 << 8;                                  // oldstyle serifs
            font.Os2Table.panose = new byte[] { 2, 11, 0, 0, 0, 0, 0, 0, 0, 0 }; // panose says sans
            Assert.AreEqual(GenericFontFamily.Serif, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        public void Classifier_NonSerifFamilyClass_FallsThroughToPanose()
        {
            // Only classes 1-7 decide. Any other class leaves the decision to PANOSE.
            var font = LoadPrivateRoboto();
            font.Os2Table.sFamilyClass = 8 << 8;                                 // sans serif class
            font.Os2Table.panose = new byte[] { 2, 2, 0, 0, 0, 0, 0, 0, 0, 0 }; // panose says serif
            Assert.AreEqual(GenericFontFamily.Serif, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        [DataRow((byte)2)]
        [DataRow((byte)6)]
        [DataRow((byte)10)]
        public void Classifier_PanoseSerifStyle_IsSerif(byte serifStyle)
        {
            var font = LoadPrivateRoboto();
            font.Os2Table.panose = new byte[] { 2, serifStyle, 0, 0, 0, 0, 0, 0, 0, 0 };
            Assert.AreEqual(GenericFontFamily.Serif, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        [DataRow((byte)0)]  // any
        [DataRow((byte)1)]  // no fit
        [DataRow((byte)11)] // normal sans
        [DataRow((byte)15)] // rounded
        public void Classifier_PanoseOtherSerifStyle_IsSansSerif(byte serifStyle)
        {
            var font = LoadPrivateRoboto();
            font.Os2Table.panose = new byte[] { 2, serifStyle, 0, 0, 0, 0, 0, 0, 0, 0 };
            Assert.AreEqual(GenericFontFamily.SansSerif, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        public void Classifier_PanoseMonospacedProportion_IsMonospace()
        {
            var font = LoadPrivateRoboto();
            font.Os2Table.panose = new byte[] { 2, 11, 0, 9, 0, 0, 0, 0, 0, 0 };
            Assert.AreEqual(GenericFontFamily.Monospace, GenericFontFamilyClassifier.Classify(font));
        }

        [TestMethod]
        public void Classifier_PanoseNotLatinText_IsIgnored()
        {
            var font = LoadPrivateRoboto();
            font.Os2Table.panose = new byte[] { 3, 2, 0, 9, 0, 0, 0, 0, 0, 0 }; // family type 3 = script
            Assert.AreEqual(GenericFontFamily.SansSerif, GenericFontFamilyClassifier.Classify(font));
        }
    }
}
