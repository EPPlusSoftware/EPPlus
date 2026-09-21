using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.TextLayout
{
    [TestClass]
    public class WrapTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void WrapMultipleFragments_SpacedEndWord()
        {

            RequireFont(SystemFontsEngine, "Aptos Narrow", FontSubFamily.Regular);
            RequireFont(SystemFontsEngine, "Goudy Stout", FontSubFamily.Regular);

            List<string> txtRuns =
            [
                "H",
                "IJ",
                "K",
                "L",
                "M ",
                "NOPE",
            ];


            var mf = new MeasurementFont();
            mf.FontFamily = "Aptos Narrow";
            mf.Style = MeasurementFontStyles.Regular;
            mf.Size = 16;

            var mf2 = new MeasurementFont();
            mf2.FontFamily = "Goudy Stout";
            mf2.Style = MeasurementFontStyles.Regular;
            mf2.Size = 11;

            List<MeasurementFont> fonts =
            [
                mf,
                mf,
                mf,
                mf,
                mf
            ];

            fonts.Add(mf2);

            var txtMeasurer = SystemFontsEngine.GetTextLayoutEngineForFont(mf2);
            var maxWidth = 114d;

            var wrappedFragments = txtMeasurer.WrapRichText(txtRuns, fonts, maxWidth.PixelToPoint());

            Assert.AreEqual(2, wrappedFragments.Count);
            Assert.AreEqual("HIJKLM", wrappedFragments[0]);
            Assert.AreEqual("NOPE", wrappedFragments[1]);
        }

        [TestMethod]
        public void WrapMultipleFragments_LongPlusEndWord()
        {
            RequireFont(SystemFontsEngine, "Aptos Narrow", FontSubFamily.Regular);
            List<string> txtRuns =
            [
                "H",
                "IJ",
                "K",
                "L",
                "Mpqrstvdef",
                " ",
                "NOPE",
            ];


            var mf = new MeasurementFont();
            mf.FontFamily = "Aptos Narrow";
            mf.Style = MeasurementFontStyles.Regular;
            mf.Size = 16;

            var mf2 = new MeasurementFont();
            mf2.FontFamily = "Aptos Narrow";
            mf2.Style = MeasurementFontStyles.Regular;
            mf2.Size = 11;

            List<MeasurementFont> fonts =
            [
                mf,
                mf,
                mf,
                mf,
                mf
            ];

            fonts.Add(mf2);
            fonts.Add(mf2);

            var txtMeasurer = SystemFontsEngine.GetTextLayoutEngineForFont(mf);



            var maxWidth = 114d;

            var wrappedFragments = txtMeasurer.WrapRichText(txtRuns, fonts, maxWidth.PixelToPoint());

            Assert.AreEqual(2, wrappedFragments.Count);
            Assert.AreEqual("HIJKLMpqrst", wrappedFragments[0]);
            Assert.AreEqual("vdef NOPE", wrappedFragments[1]);
        }
    }
}
