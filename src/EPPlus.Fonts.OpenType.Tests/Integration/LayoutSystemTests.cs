using EPPlus.Fonts.OpenType.Integration;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Fonts.OpenType.Integration.RichText;
using EPPlus.Fonts.OpenType.Utils;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.RichText;
using OfficeOpenXml.Style;
using System.Drawing;

namespace EPPlus.Fonts.OpenType.Tests.Integration
{
    [TestClass]
    public class LayoutSystemTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void TestParagraphs()
        {

            List<string> lstOfRichText = new() { "MyparticularilyLongWord", "WithAbsolutelyNoSpacesAtAllJustToBeDifficult" };
            var font = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };
            var font2 = new MeasurementFont()
            {
                FontFamily = "Goudy Stout",
                Size = 16,
                Style = MeasurementFontStyles.Regular
            };

            List<MeasurementFont> fonts = new List<MeasurementFont>() { font, font2 };

            var fragments = new List<ITextFragmentBase>();

            for (int i = 0; i < lstOfRichText.Count(); i++)
            {
                var currentFrag = new TextFragment() { Text = lstOfRichText[i], Font = fonts[i] };
                fragments.Add(currentFrag);
            }

            var paragraph = new LayoutSystem(fragments);
            var styleRuns = paragraph.GetTextOfAllTextRuns();

            Assert.AreEqual(lstOfRichText[0], styleRuns[0]);
            Assert.AreEqual(lstOfRichText[1], styleRuns[1]);


            var layout = OpenTypeFonts.GetTextLayoutEngineForFont(font);
            var wrappedLines = layout.WrapRichTextLines(fragments, 225d);

            var wrappedLinesPara = paragraph.Wrap(225d);

            Assert.AreEqual(wrappedLines.Count, wrappedLinesPara.Count);

            for (int i = 0; i < wrappedLines.Count; i++)
            {
                Assert.AreEqual(wrappedLines[i].Text, wrappedLinesPara[i].Text);
                Assert.AreEqual(wrappedLines[i].Width, wrappedLinesPara[i].Width);
            }
        }

        [TestMethod]
        public void TestLayoutSystemParagraphChars()
        {
            List<string> lstOfRichText = new() { "Here comes lorem ipsum\u2029 " +
                "Sed ut perspiciatis, unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam eaque ipsa, quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt, explicabo. Nemo enim ipsam voluptatem, quia voluptas sit, aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos, qui ratione voluptatem sequi nesciunt, neque porro quisquam est, qui dolorem ipsum, quia dolor sit amet consectetur adipisci[ng] velit, sed quia non numquam [do] eius modi tempora inci[di]dunt, ut labore et dolore magnam aliquam quaerat voluptatem. Ut enim ad minima veniam, quis nostrum[d] exercitationem ullam corporis suscipit laboriosam, nisi ut aliquid ex ea commodi consequatur? [D]Quis autem vel eum i[r]ure reprehenderit, qui in ea voluptate velit esse, quam nihil molestiae consequatur, vel illum, qui dolorem eum fugiat, quo voluptas nulla pariatur?\u2029 " +
                "At vero eos et accusamus et iusto odio dignissimos ducimus, qui blanditiis praesentium voluptatum deleniti atque corrupti, quos dolores et quas molestias excepturi sint, obcaecati cupiditate non provident, similique sunt in culpa, qui officia deserunt mollitia animi, id est laborum et dolorum fuga. Et harum quidem reru[d]um facilis est e[r]t expedita distinctio. Nam libero tempore, cum soluta nobis est eligendi optio, cumque nihil impedit, quo minus id, quod maxime placeat facere possimus, omnis voluptas assumenda est, omnis dolor repellend[a]us. Temporibus autem quibusdam et aut officiis debitis aut rerum necessitatibus saepe eveniet, ut et voluptates repudiandae sint et molestiae non recusandae. Itaque earum rerum hic tenetur a sapiente delectus, ut aut reiciendis voluptatibus maiores alias consequatur aut perferendis doloribus asperiores repellat.\u2029 " +
                "Let's see if we can recognize unicode paragraph separators" };
            var font = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };
            var fragments = new List<ITextFragmentBase>()
            {
                new TextFragment() {Text = lstOfRichText[0], Font = font }
            };

            var layout = new LayoutSystem(fragments);
            Assert.AreEqual(3, layout.GetParagraphSeparatorCount());
        }

        [TestMethod]
        public void TestParagraphs_DifficultCase()
        {
            List<string> lstOfRichText = new() { "TextBox2", "ra underline", "La Strike", "Goudy size 16" };
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


            List<MeasurementFont> fonts = new() { font2, font3, font4, font5 };
            var fragments = new List<ITextFragmentBase>();

            for (int i = 0; i < lstOfRichText.Count(); i++)
            {
                var currentFrag = new TextFragment() { Text = lstOfRichText[i], Font = fonts[i] };
                fragments.Add(currentFrag);
            }

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();

            var paragraph = new LayoutSystem(fragments);
            var wrappedLines = paragraph.Wrap(225d);

            var line1 = wrappedLines[0];
        }

        [TestMethod]
        public void EnsureCorrectTotalIndex()
        {
            List<string> lstOfRichText = new() { "aaaaaaaa aa aaaaaaaaaLa Strike", "Goudy size 16" };
            var font = new MeasurementFont()
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Bold
            };
            var font2 = new MeasurementFont()
            {
                FontFamily = "Goudy Stout",
                Size = 16,
                Style = MeasurementFontStyles.Regular
            };

            List<MeasurementFont> fonts = new List<MeasurementFont>() { font, font2 };

            var fragments = new List<ITextFragmentBase>();

            for (int i = 0; i < lstOfRichText.Count(); i++)
            {
                var currentFrag = new TextFragment() { Text = lstOfRichText[i], Font = fonts[i] };
                fragments.Add(currentFrag);
            }

            var paragraph = new LayoutSystem(fragments);
            var wrappedLines = paragraph.Wrap(225d);

            Assert.AreEqual("StrikeGoudy size", wrappedLines[1].Text);
            Assert.AreEqual(24, wrappedLines[1].LineFragments[0].StartFullTextIdx);
            Assert.AreEqual(24, wrappedLines[1].LineFragments[0].StartRtIdx);
        }

        [TestMethod]
        public void EnsureRTCharIdxBecomesCorrectWhenBreaking()
        {
            List<string> lstOfRichText = new() { "MyparticularilyLongWord", "WithAbsolutelyNoSpacesAtAllJustToBeDifficult" };
            var font = new MeasurementFont()
            {
                FontFamily = "Archivo Narrow",
                Size = 12,
                Style = MeasurementFontStyles.Regular
            };
            var font2 = new MeasurementFont()
            {
                FontFamily = "Oi",
                Size = 20,
                Style = MeasurementFontStyles.Regular
            };

            List<MeasurementFont> fonts = new List<MeasurementFont>() { font, font2 };

            var fragments = new List<ITextFragmentBase>();

            for (int i = 0; i < lstOfRichText.Count(); i++)
            {
                var currentFrag = new TextFragment() { Text = lstOfRichText[i], Font = fonts[i] };
                fragments.Add(currentFrag);
            }

            var shaper = OpenTypeFonts.GetShaperForFont(font2);
            //var shapes = shaper.ShapeLight("WithAbsolutelyNoSpacesAtAllJustToBeDifficult");
            var layout = new TextLayoutEngine(shaper);
            var wrappedLines = layout.WrapRichTextLines(fragments, 225d);

            var paragraph = new LayoutSystem(fragments);
            var wrappedLines2 = paragraph.Wrap(225d);
            //var layout = OpenTypeFonts.GetTextLayoutEngineForFont(font, FontFolders);
            //var wrappedLines = layout.WrapRichTextLines(fragments, 225d);

            Assert.AreEqual(5, wrappedLines[1].LineFragments[0].StartRtIdx);
            Assert.AreEqual(17, wrappedLines[2].LineFragments[0].StartRtIdx);
            Assert.AreEqual(29, wrappedLines[3].LineFragments[0].StartRtIdx);
            Assert.AreEqual(41, wrappedLines[4].LineFragments[0].StartRtIdx);

            Assert.AreEqual(5, wrappedLines2[1].LineFragments[0].StartRtIdx);
            Assert.AreEqual(17, wrappedLines2[2].LineFragments[0].StartRtIdx);
            Assert.AreEqual(29, wrappedLines2[3].LineFragments[0].StartRtIdx);
            Assert.AreEqual(41, wrappedLines2[4].LineFragments[0].StartRtIdx);
        }


        [TestMethod]
        public void EnsureWrappingSimplePlainTextCorrectly()
        {
            string myText = "Hi! I am a simple but somewhat wordy text string that is being Tested for wrapping in the case where only one font exists";
            var font = new MeasurementFont()
            {
                FontFamily = "Archivo Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Regular
            };

            List<MeasurementFont> fonts = new List<MeasurementFont>() { font };
            List<string> strings = new List<string>() { myText };


            var fragments = new List<ITextFragmentBase>();

            for (int i = 0; i < strings.Count(); i++)
            {
                var currentFrag = new TextFragment() { Text = strings[i], Font = fonts[i] };
                fragments.Add(currentFrag);
            }

            var paragraph = new LayoutSystem(fragments);

            var lines = paragraph.Wrap(92.976377953d);

            Assert.AreEqual("Hi! I am a simple but", lines[0].Text);
            Assert.AreEqual("somewhat wordy text", lines[1].Text);
            Assert.AreEqual("string that is being", lines[2].Text);
            Assert.AreEqual("Tested for wrapping in", lines[3].Text);
            Assert.AreEqual("the case where only", lines[4].Text);
            Assert.AreEqual("one font exists", lines[5].Text);
        }

        [TestMethod]
        public void EnsureWrappingRichTextAndGettingLineSpacing()
        {
            List<string> txtLst = new List<string>() { "Hi ", "I am rich ", "But I am Even Richer " };
            var rt = new OpenTypeRichTextBase(txtLst[0], "Roboto", 12f);
            var rtSecond = new OpenTypeRichTextBase(txtLst[1], "Archivo Narrow", 11f);
            var rtThird = new OpenTypeRichTextBase(txtLst[2], "Roboto", 18f);

            rtThird.Italic = true;
            rtThird.Bold = true;

            List<IRichTextFormatBase> rtLst = new List<IRichTextFormatBase>() { rt, rtSecond, rtThird};

            var paragraph = new LayoutSystem(rtLst);

            var lines = paragraph.Wrap(92.976377953d);

            //Assert correct wrapping
            Assert.AreEqual("Hi I am rich But I", lines[0].Text);
            Assert.AreEqual("am Even", lines[1].Text);
            Assert.AreEqual("Richer", lines[2].Text);

            //Assert line segments correct count
            Assert.AreEqual(3, lines[0].LineFragments.Count);
            Assert.AreEqual(3, lines[0].InternalLineFragments.Count);
            Assert.AreEqual(1, lines[1].LineFragments.Count);
            Assert.AreEqual(1, lines[1].InternalLineFragments.Count);
            Assert.AreEqual(1, lines[2].LineFragments.Count);
            Assert.AreEqual(1, lines[2].InternalLineFragments.Count);

            //Assert correct fragment in correct spot
            Assert.AreEqual(rtLst[0], (IRichTextFormatBase)lines[0].LineFragments[0].OriginalTextFragment.RichTextFormat);
            Assert.AreEqual(rtLst[1], (IRichTextFormatBase)lines[0].LineFragments[1].OriginalTextFragment.RichTextFormat);
            Assert.AreEqual(rtLst[2], (IRichTextFormatBase)lines[0].LineFragments[2].OriginalTextFragment.RichTextFormat);

            Assert.AreEqual(rtLst[2], (IRichTextFormatBase)lines[1].LineFragments[0].OriginalTextFragment.RichTextFormat);
            Assert.AreEqual(rtLst[2], (IRichTextFormatBase)lines[2].LineFragments[0].OriginalTextFragment.RichTextFormat);
        }

        [TestMethod]
        public void TestSimpleRichText()
        {
            //var rtCollection = new RichTextCollectionBase();
            //var someTextRt = rtCollection.Add("SomeText", true);
            //var richRt = rtCollection.Add("rich");
            //var richerRt = rtCollection.Add("richer");
            //var richestRt = rtCollection.Add("richest");
            //var wealthyRt = rtCollection.Add("Wealthy");

            //richRt.Info.FontFamily = "Roboto";
            //richRt.Info.Italic = true;

            //richerRt.Info.FontFamily = "Roboto";
            //richerRt.Info.Size = 16;
            //richerRt.Info.Italic = true;
            //richerRt.Info.UnderlineType = (int)ExcelUnderLineType.Single;

            //richestRt.Info.FontFamily = "Oi";
            //richestRt.Info.FontColor = Color.BlueViolet;
            //richestRt.Info.Bold = true;
            //richestRt.Info.Italic = true;
            //richerRt.Info.Size = 18;


            //someTextRt.FontData.Family = "Archivo Narrow";

            //richRt
        }
    }
}
