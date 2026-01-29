using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer;
using EPPlus.Fonts.OpenType.TrueTypeMeasurer.DataHolders;
using EPPlus.Fonts.OpenType.Utils;
using EPPlus.Graphics;
using EPPlusImageRenderer;
using EPPlusImageRenderer.Svg;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Globalization;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Tests
{
    [TestClass]
    public sealed class TextRenderTests : TestBase
    {
        private void SetFillColor(ExcelDrawingFillBasic fill, string color)
        {
            if (color.StartsWith("#"))
            {
                fill.Style = eFillStyle.SolidFill;
                var r = int.Parse(color.Substring(1, 2), NumberStyles.HexNumber);
                var g = int.Parse(color.Substring(3, 2), NumberStyles.HexNumber);
                var b = int.Parse(color.Substring(5, 2), NumberStyles.HexNumber);
                var c = System.Drawing.Color.FromArgb(0xFF, (byte)r, (byte)g, (byte)b);
                fill.SolidFill.Color.SetRgbColor(c);
            }
            else if (string.IsNullOrWhiteSpace(color) == false)
            {
                fill.Style = eFillStyle.SolidFill;
                try
                {
                    var c = System.Drawing.Color.FromName(color);
                    if (c.IsEmpty)
                    {
                        var sc = Enum.Parse<eSchemeColor>(color);
                        fill.SolidFill.Color.SetSchemeColor(sc);
                    }
                    else
                    {
                        fill.SolidFill.Color.SetPresetColor(c);
                    }
                }
                catch
                {
                    var sc = Enum.Parse<eSchemeColor>(color);
                    fill.SolidFill.Color.SetSchemeColor(sc);
                }
            }
            else
            {
                fill.Style = eFillStyle.NoFill;
            }
        }

        TextFragmentCollection GenerateTextFragments(ExcelDrawingTextRunCollection runs)
        {
            List<string> runContents = new List<string>();
            List<float> fontSizes = new List<float>();

            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasurementFont();

                runContents.Add(txtRun.Text);
                fontSizes.Add(runFont.Size);
            }

            return new TextFragmentCollection(runContents, fontSizes);
        }

        List<TextLineSimple> GetWrappedText(ExcelDrawingTextRunCollection runs, TextFragmentCollection fragments)
        {
            FontMeasurerTrueType ttMeasurer = new();
            List<MeasurementFont> fonts = new List<MeasurementFont>();

            for (int i = 0; i < runs.Count(); i++)
            {
                var txtRun = runs[i];
                var runFont = txtRun.GetMeasurementFont();
                fonts.Add(runFont);
            }

            var maxSizePoints = Math.Round(300d, 0, MidpointRounding.AwayFromZero).PixelToPoint();
            return ttMeasurer.WrapMultipleTextFragmentsToTextLines(fragments, fonts, maxSizePoints);
        }

        [TestMethod]
        public void VerifyTextRunBounds()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var cube = ws.Drawings.AddShape("Cube1", OfficeOpenXml.Drawing.eShapeStyle.Cube);

                cube.Font.Color = System.Drawing.Color.Goldenrod;

                cube.TextBody.TopInsert = 0;
                cube.TextBody.BottomInsert = 0;
                cube.TextBody.RightInsert = 0;
                cube.TextBody.LeftInsert = 0;

                var para1 = cube.TextBody.Paragraphs.Add("TextBox\r\na");

                para1.LeftMargin = 5;
                cube.TextBody.TopInsert = 10;

                var para2 = cube.TextBody.Paragraphs.Add("TextBox2");
                para2.TextRuns[0].FontItalic = true;
                para2.TextRuns[0].FontBold = true;
                para2.TextRuns.Add("ra underline").FontUnderLine = eUnderLineType.Dash;
                para2.TextRuns.Add("La Strike").FontStrike = eStrikeType.Single;
                var tRun1 = para2.TextRuns.Add("Goudy size 16");
                tRun1.SetFromFont("Goudy Stout", 16);
                tRun1.Fill.Color = System.Drawing.Color.IndianRed;
                var tRun2 = para2.TextRuns.Add("SvgSize 24");
                tRun2.FontSize = 24;

                cube.TextAlignment = eTextAlignment.Left;
                cube.TextAnchoring = eTextAnchoringType.Top;

                cube.TextBody.HorizontalTextOverflow = eTextHorizontalOverflow.Clip;
                cube.TextBody.VerticalTextOverflow = eTextVerticalOverflow.Clip;

                SetFillColor(cube.Fill, "#156082");
                SetFillColor(cube.Border.Fill, "#042433");

                var aFont = cube.Font;
                var paragraph0 = cube.TextBody.Paragraphs[0];

                var autofit = cube.TextBody.TextAutofit;

                cube.GetSizeInPixels(out int testWidth, out int testHeight);

                var parentBB = new BoundingBox();
                parentBB.Width = testWidth;
                parentBB.Height = testHeight;

                var svgShape = new SvgShape(cube);
                SvgTextBodyItem tbItem = new SvgTextBodyItem(svgShape, parentBB);

                //Verify new method works at all
                for(int i = 1; i< cube.TextBody.Paragraphs.Count; i++)
                {
                    var fragments = GenerateTextFragments(cube.TextBody.Paragraphs[i].TextRuns);

                    var wrappedText = GetWrappedText(cube.TextBody.Paragraphs[i].TextRuns, fragments);
                    if(i == 0)
                    {
                        Assert.AreEqual("TextBox", wrappedText[0].Text);
                        Assert.AreEqual("a", wrappedText[1].Text);
                    }
                    if(i == 1)
                    {
                        Assert.AreEqual("TextBox2ra underlineLa", wrappedText[0].Text);
                        Assert.AreEqual("StrikeGoudy size", wrappedText[1].Text);
                        Assert.AreEqual("16SvgSize 24", wrappedText[2].Text);
                    }
                }
                tbItem.ImportTextBody(cube.TextBody);

                var txtRun1Bounds = tbItem.Paragraphs[0].Runs[0].Bounds;

                Assert.AreEqual(43.835286458333336d, txtRun1Bounds.Width);

                var widthLine2 = tbItem.Paragraphs[0].Runs[0].PerLineWidth[1];
                var topYLine2 = tbItem.Paragraphs[0].Runs[0].YIncreasePerLine[1];
                Assert.AreEqual(7.147135416666667d, widthLine2);
                Assert.AreEqual(17.903645833333336d, topYLine2);

                var txtRuns2 = tbItem.Paragraphs[1].Runs;

                Assert.AreEqual(53.20963541666667d, txtRuns2[0].Bounds.Width);
                var currentLineWidth = txtRuns2[0].Bounds.Width;

                Assert.AreEqual(currentLineWidth, txtRuns2[1].Bounds.Left);
                Assert.AreEqual(69.55924479166667d, txtRuns2[1].Bounds.Width);
                currentLineWidth += txtRuns2[1].Bounds.Width;

                Assert.AreEqual(currentLineWidth, txtRuns2[2].Bounds.Left);
                Assert.AreEqual(49.89388020833334, txtRuns2[2].Bounds.Width);
                currentLineWidth += txtRuns2[2].Bounds.Width;

                Assert.AreEqual(21.333333333333332, txtRuns2[3].Bounds.Height);
                Assert.AreEqual(currentLineWidth, txtRuns2[3].Bounds.Left);
                Assert.AreEqual(283.21875d, txtRuns2[3].Bounds.Width);
                currentLineWidth += txtRuns2[3].Bounds.Width;

                Assert.AreEqual(32, txtRuns2[4].Bounds.Height);
                Assert.AreEqual(currentLineWidth, txtRuns2[4].Bounds.Left);
                Assert.AreEqual(134.20312500000006d, txtRuns2[4].Bounds.Width);
                currentLineWidth += txtRuns2[4].Bounds.Width;
            }
        }

        [TestMethod]
        public void ConceptLines()
        {
            var currentChar = 'a';
            var defaultFont = new MeasurementFont
            {
                FontFamily = "Aptos Narrow",
                Size = 11,
                Style = MeasurementFontStyles.Regular
            };

            List<string> manyRichText = new List<string>();
            List<MeasurementFont> manyFonts = new List<MeasurementFont>();

            for (int i = 0; i< 20; i++)
            {
                manyRichText.Add(currentChar.ToString());
                manyFonts.Add(defaultFont);
                currentChar++;
                defaultFont.Size ++;
            }

            var fragments = new TextFragmentCollection(manyRichText);
            var fontMeasurer = new FontMeasurerTrueType();

            var strings = fontMeasurer.WrapMultipleTextFragments(fragments, manyFonts, 30d.PixelToPoint());

            var outputLines = fragments.GetOutputLines();


            //var paragraph = new TextParagraph(fragments, manyFonts);

            //List<string> manyRichText = new List<string>() {"a","b","c","d","e","f","g" };
        }

        [TestMethod]
        public void ConceptTextRun()
        {
            string rt1 = "My richtext1\r\n of len";
            string rt2 = "gth beyond and then I am richtext2";

            string combined = rt1 + rt2;

            int charMax = 10;

            int lineCharCount = 0;

            List<string> lines = new List<string>();

            for (int i = 0; i < combined.Length; i++)
            {
                if(lineCharCount > charMax)
                {
                    var currLine = combined.Substring(i - lineCharCount, lineCharCount);
                    lines.Add(currLine);
                    lineCharCount = 0;
                }

                if (combined[i] == '\r')
                {
                    var currLine = combined.Substring(i - lineCharCount, lineCharCount);
                    lines.Add(currLine);
                    lineCharCount = 0;
                    i++;
                    continue;
                }

                lineCharCount++;
            }

            var finalLine = combined.Substring(combined.Length - lineCharCount, lineCharCount);
            lines.Add(finalLine);

            foreach (string line in lines)
            {
                Debug.WriteLine(line);
            }

        }
    }
}
