using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using EPPlus.Graphics;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Engineering;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System.Drawing;
using System.Text;
using static OfficeOpenXml.Drawing.OleObject.Structures.OleObjectDataStructures;

namespace EPPlus.Export.ImageRenderer.Tests.DrawingShapeRenderer
{
    [TestClass]
    public class SvgStandAloneTests : TestBase
    {

        private GroupRenderItem GenerateShapeRenderer()
        {
            BoundingBox bounds = new BoundingBox(0, 0, 500, 500);
            StringBuilder sb = new StringBuilder();
            var svgShapeRenderer = new SvgShapeRenderer(bounds, sb);


            var baseGroup = new GroupRenderItem(bounds);

            var background = new RectRenderItem(baseGroup.Bounds);

            background.Width = bounds.Width;
            background.Height = bounds.Height;
            background.FillColor = "aliceBlue";

            baseGroup.AddChildItem(background);
            return baseGroup;
        }

        private GroupRenderItem GenerateGroupRenderItem()
        {
            BoundingBox bounds = new BoundingBox(0, 0, 500, 500);

            var baseGroup = new GroupRenderItem(bounds);

            var background = new RectRenderItem(baseGroup.Bounds);

            background.Width = bounds.Width;
            background.Height = bounds.Height;
            background.FillColor = "aliceBlue";

            baseGroup.AddChildItem(background);
            return baseGroup;
        }

        private void GenerateSvgFile(string fileName, BoundingBox bounds, params RenderItem[] items)
        {

            StringBuilder sb = new StringBuilder();
            var svgShapeRenderer = new SvgShapeRenderer(bounds, sb);

            List<RenderItem> renderItems = items.ToList();
            svgShapeRenderer.Render(renderItems);

            var svg = sb.ToString();

            SaveTextFileToWorkbook($"svg\\{fileName}.svg", svg);
        }

        [TestMethod]
        public void SvgRectTest()
        {
            var baseGroup = GenerateShapeRenderer();
            GenerateSvgFile("rectStandAlone", baseGroup.Bounds, baseGroup);
        }

        [TestMethod]
        public void SvgTextRun()
        {
            var baseGroup = GenerateShapeRenderer();

            var rt = new RichTextFormatSimple();
            rt.Text = "My text";
            rt.UnderlineType = 1;
            rt.FontColor = System.Drawing.Color.Black;
            rt.Family = "Archivo Narrow";
            rt.SubFamily = OfficeOpenXml.Interfaces.Fonts.FontSubFamily.Regular;
            rt.Size = 12f;

            var textRun = new SvgTextRunRenderItem(baseGroup.Bounds, rt, rt.Text, true);

            //Add size of text since svg renders text upwards from the start point.
            textRun.YPosition = rt.Size;

            baseGroup.AddChildItem(textRun);

            GenerateSvgFile("textRunStandAlone", baseGroup.Bounds, baseGroup);
        }

        private void GenerateTextBodyFile(string fileName, GroupRenderItem baseGroup, SvgTextBodyRenderItem textBody)
        {
            StringBuilder sb = new StringBuilder();
            var svgShapeRenderer = new SvgShapeRenderer(baseGroup.Bounds, sb);

            var background = new RectRenderItem(baseGroup.Bounds);

            background.Width = baseGroup.Bounds.Width;
            background.Height = baseGroup.Bounds.Height;
            background.FillColor = "aliceBlue";

            baseGroup.AddChildItem(textBody);
            baseGroup.AddChildItem(background);

            List<RenderItem> items = new List<RenderItem>() { baseGroup };

            svgShapeRenderer.Render(items);

            var svg = sb.ToString();


            SaveTextFileToWorkbook($"svg\\{fileName}.svg", svg);
        }


        private SvgTextBodyRenderItem GenerateTextBody(GroupRenderItem baseGroup)
        {
            var textBody = new SvgTextBodyRenderItem(baseGroup.Bounds, true);
            var paragraph = textBody.AddParagraph("Hello");

            paragraph.AddText(" There");

            var rtItem = new RichTextFormatSimple("Second paragraph", "Archivo Narrow", 16f, true);
            rtItem.FontColor = Color.DarkGreen;
            var para2 = textBody.AddParagraph(rtItem);

            textBody.AddChildItem(paragraph);
            textBody.AddChildItem(para2);

            baseGroup.AddChildItem(textBody);

            return textBody;
        }

        [TestMethod]
        public void SvgTextBodyTest()
        {
            var baseGroup = GenerateGroupRenderItem();
            var textBody = GenerateTextBody(baseGroup);
            GenerateSvgFile("standAloneTextBody", baseGroup.Bounds, baseGroup);
        }

        [TestMethod]
        public void SvgTextBodyTestCenterAlignmentGenerated()
        {
            var baseGroup = GenerateGroupRenderItem();

            var textBody = GenerateTextBody(baseGroup);
            textBody.Paragraphs[0].AddText(" a\r\n new day beckons");
            textBody.Paragraphs[0].HorizontalAlignment = RenderItems.Shared.TextAlignment.Center;

            textBody.Paragraphs[1].HorizontalAlignment = RenderItems.Shared.TextAlignment.Center;
            textBody.Paragraphs[1].AddText("\r\n What fun, what fun!");

            //Text was added to the paragraph above the last paragraph
            //We must re-calculate where the next paragraph should be placed
            textBody.RecalculateParagraphs();
            textBody.ApplyAutoSize();

            double delta = 0.001;

            //new day beckons is the largest line in the centered paragraph[0]
            //Assert that the first line has been centered appropriately
            Assert.AreEqual(9.890869140625d, textBody.Paragraphs[0].Runs[0].Bounds.Left, delta);
            Assert.AreEqual(33.142333984375d, textBody.Paragraphs[0].Runs[1].Bounds.Left, delta);
            Assert.AreEqual(59.509033203125d, textBody.Paragraphs[0].Runs[2].Bounds.Left, delta);
            Assert.AreEqual(0d, textBody.Paragraphs[0].Runs[3].Bounds.Left);

            Assert.AreEqual(4.3760001659393311d, textBody.Paragraphs[1].Runs[0].Bounds.Left, delta);
            Assert.AreEqual(0d, textBody.Paragraphs[1].Runs[1].Bounds.Left);

            //Assert that the second paragraph has been moved correctly
            Assert.AreEqual(26.85546875d, textBody.Paragraphs[1].Bounds.Top);
            GenerateSvgFile("textBodyAlignCenter", baseGroup.Bounds, baseGroup);
        }

        [TestMethod]
        public void SvgTextBodyTestRightAlignmentGenerated()
        {
            var baseGroup = GenerateGroupRenderItem();

            var textBody = GenerateTextBody(baseGroup);
            textBody.Paragraphs[0].AddText(" a\r\n new day beckons");
            textBody.Paragraphs[0].HorizontalAlignment = RenderItems.Shared.TextAlignment.Right;


            textBody.Paragraphs[1].HorizontalAlignment = RenderItems.Shared.TextAlignment.Right;
            textBody.Paragraphs[1].AddText("\r\n What fun, what fun!");

            //Text was added to the paragraph above the last paragraph
            //We must re-calculate where the next paragraph should be placed
            textBody.RecalculateParagraphs();
            textBody.ApplyAutoSize();

            double delta = 0.001;

            //Assert that the first line has been aligned correctly
            Assert.AreEqual(19.78173828125d, textBody.Paragraphs[0].Runs[0].Bounds.Left, delta);
            Assert.AreEqual(43.033203125d, textBody.Paragraphs[0].Runs[1].Bounds.Left, delta);
            Assert.AreEqual(69.39990234375d, textBody.Paragraphs[0].Runs[2].Bounds.Left, delta);
            Assert.AreEqual(0d, textBody.Paragraphs[0].Runs[3].Bounds.Left);

            Assert.AreEqual(8.7520003318786621d, textBody.Paragraphs[1].Runs[0].Bounds.Left, delta);
            Assert.AreEqual(0d, textBody.Paragraphs[1].Runs[1].Bounds.Left);

            GenerateSvgFile("textBodyAlignRight", baseGroup.Bounds, baseGroup);
        }

        [TestMethod]
        public void SvgTextBodyVerticalAlignmentGenerated()
        {
            var baseGroup = GenerateGroupRenderItem();

            var textBody = GenerateTextBody(baseGroup);
            textBody.Paragraphs[0].AddText(" a\r\n new day beckons");
            textBody.Paragraphs[1].AddText("\r\n What fun, what fun!");

            //Text was added to the paragraph above the last paragraph
            //We must re-calculate where the next paragraph should be placed
            textBody.RecalculateParagraphs();
            textBody.ApplyAutoSize();

            textBody.AutoSize = false;
            textBody.Height = 500;

            textBody.Bounds.Top = 0;
            textBody.VerticalAlignment = TextAnchoringType.Center;
            textBody.Bounds.Top = textBody.GetAlignmentVertical();

            double delta = 0.001;

            Assert.AreEqual(180.04052829742432d, textBody.Bounds.Top, delta);

            GenerateSvgFile("textBodyAlignVCenter", baseGroup.Bounds, baseGroup);
        }

        [TestMethod]
        public void SvgTextBodyVerticalAlignmentBottomGenerated()
        {
            var baseGroup = GenerateGroupRenderItem();

            var textBody = GenerateTextBody(baseGroup);
            textBody.Paragraphs[0].AddText(" a\r\n new day beckons");
            textBody.Paragraphs[1].AddText("\r\n What fun, what fun!");

            //Text was added to the paragraph above the last paragraph
            //We must re-calculate where the next paragraph should be placed
            textBody.RecalculateParagraphs();
            textBody.ApplyAutoSize();

            textBody.AutoSize = false;
            textBody.Height = 500;

            textBody.Bounds.Top = 0;
            textBody.VerticalAlignment = TextAnchoringType.Bottom;
            textBody.Bounds.Top = textBody.GetAlignmentVertical();

            double delta = 0.001;
            Assert.AreEqual(430.04052829742432d, textBody.Bounds.Top, delta);

            GenerateSvgFile("textBodyAlignVBottom", baseGroup.Bounds, baseGroup);
        }

        private RenderTextbox GenerateTextBox(out GroupRenderItem group)
        {
            group = GenerateGroupRenderItem();

            var textbox = new RenderTextbox(group.Bounds, 500d, 500d);
            textbox.TextBody = new SvgTextBodyRenderItem(group.Bounds, true);
            var paragraph = textbox.TextBody.AddParagraph("Hello");

            paragraph.AddText(" There");

            var rtItem = new RichTextFormatSimple("Second paragraph", "Archivo Narrow", 16f, true);
            rtItem.FontColor = Color.DarkGreen;
            var para2 = textbox.TextBody.AddParagraph(rtItem);

            textbox.Rectangle.FillColor = "#F9F6C4";

            return textbox;
        }

        [TestMethod]
        public void BasicTextBox()
        {
            var textbox = GenerateTextBox(out GroupRenderItem group);
            textbox.AppendRenderItems(group.RenderItems);

            double delta = 0.001;

            Assert.AreEqual(107.95200681686401d, textbox.Width, delta);
            Assert.AreEqual(34.979735851287842d, textbox.Height, delta);

            GenerateSvgFile("BasicTextBox", group.Bounds, group);
        }

        [TestMethod]
        public void TextBoxWithMargins()
        {
            var textbox = GenerateTextBox(out GroupRenderItem group);

            double delta = 0.001;

            textbox.LeftMargin = 10d;
            textbox.TopMargin = 10d;

            textbox.AppendRenderItems(group.RenderItems);

            //Assert local position unchanged
            Assert.AreEqual(0d, textbox.TextBody.Left);
            Assert.AreEqual(0d, textbox.TextBody.Top);

            //Assert global position changed
            Assert.AreEqual(10d, textbox.TextBody.Bounds.Position.X);
            Assert.AreEqual(10d, textbox.TextBody.Bounds.Position.Y);

            //Assert width and height changed by margins
            Assert.AreEqual(117.95200681686401d, textbox.Width, delta);
            Assert.AreEqual(44.979735851287842d, textbox.Height, delta);


            GenerateSvgFile("MarginTextBox", group.Bounds, group);
        }

        [TestMethod]
        public void TextBoxWithAllMargins()
        {
            var textbox = GenerateTextBox(out GroupRenderItem group);

            double delta = 0.001;

            textbox.LeftMargin = 10d;
            textbox.TopMargin = 10d;
            textbox.RightMargin = 10d;
            textbox.BottomMargin = 10d;

            textbox.AppendRenderItems(group.RenderItems);

            //Assert width and height changed by margins
            Assert.AreEqual(127.95200681686401d, textbox.Width, delta);
            Assert.AreEqual(54.979735851287842d, textbox.Height, delta);

            GenerateSvgFile("AllMarginsTextBox", group.Bounds, group);
        }


        /// <summary>
        /// TODO: Discuss. Should it really work like this?
        /// There IS an argument to be made that margin should BE textbody position
        /// At the same time then positioning in accordance with vertical aligment then becomes difficult
        /// And might affect the margin
        /// </summary>
        [TestMethod]
        public void TextBoxWithAllMarginsANDTextbodyChanged()
        {
            var textbox = GenerateTextBox(out GroupRenderItem group);

            double delta = 0.001;

            textbox.TextBody.AutoSize = true;

            textbox.TextBody.Left = 15d;
            textbox.TextBody.Top = 15d;

            textbox.LeftMargin = 10d;
            textbox.TopMargin = 10d;
            textbox.RightMargin = 10d;
            textbox.BottomMargin = 10d;

            textbox.AppendRenderItems(group.RenderItems);

            //Assert width and height changed by margins and textbody
            Assert.AreEqual(142.952006816864d, textbox.Width, delta);
            Assert.AreEqual(69.979735851287842d, textbox.Height, delta);

            GenerateSvgFile("TextAnchor_TextBox", group.Bounds, group);
        }
    }
}
