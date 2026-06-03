using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Graphics;
using OfficeOpenXml.Drawing.Renderer.TextBox;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EPPlus.DrawingRenderer.RenderItems.SvgItem;
using EPPlus.Fonts.OpenType.Integration.RichText;
using EPPlus.Fonts.OpenType.Integration.DataHolders;
using System.Drawing;

namespace EPPlus.Export.ImageRenderer.Tests.DrawingShapeRenderer
{
    [TestClass]
    public class SvgStandAloneTests : TestBase
    {
        [TestMethod]
        public void SvgRectTest()
        {
            BoundingBox bounds = new BoundingBox(0,0,500,500);
            StringBuilder sb = new StringBuilder();
            var svgShapeRenderer = new SvgShapeRenderer(bounds, sb);

            var baseGroup = new GroupRenderItem(bounds);

            var rectItem = new RectRenderItem(baseGroup.Bounds);

            rectItem.Width = 250;
            rectItem.Height = 250;
            rectItem.FillColor = "darkblue";

            //var textBody = new RenderTextBody(baseGroup.Bounds, true);

            //textBody.Text = "Hello";
            //var para = new SvgParagraphRenderItem(textBody, textBody.Bounds);
            
            //var para2 = new DrawingParagraphRenderItem(textBody, textBody.Bounds);
            //textBody.Paragraphs.Add

            baseGroup.AddChildItem(rectItem);
            //baseGroup.AddChildItem(textBody);

            List<RenderItem> items = new List<RenderItem>() { baseGroup };

            svgShapeRenderer.Render(items);

            var svg = sb.ToString();

            SaveTextFileToWorkbook("svg\\rectStandalone.svg", svg);
        }

        [TestMethod]
        public void SvgTextRun()
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

            var rt = new RichTextFormatSimple();
            rt.Text = "My text";
            rt.UnderlineType = 1;
            rt.FontColor = System.Drawing.Color.Black;
            rt.Family = "Archivo Narrow";
            rt.SubFamily = OfficeOpenXml.Interfaces.Fonts.FontSubFamily.Regular;
            rt.Size = 12f;
           
            //var paragraph = new SvgParagraphRenderItem()

            var textRun = new SvgTextRunRenderItem(baseGroup.Bounds, rt, rt.Text);
            baseGroup.AddChildItem(textRun);


            List<RenderItem> items = new List<RenderItem>() { baseGroup };

            svgShapeRenderer.Render(items);

            var svg = sb.ToString();


            SaveTextFileToWorkbook("svg\\textRunStandAlone.svg", svg);
        }

        [TestMethod]
        public void SvgTextBodyTest()
        {
            BoundingBox bounds = new BoundingBox(0, 0, 500, 500);
            StringBuilder sb = new StringBuilder();
            var svgShapeRenderer = new SvgShapeRenderer(bounds, sb);
            

            var baseGroup = new GroupRenderItem(bounds);

            var background = new RectRenderItem(baseGroup.Bounds);

            background.Width = bounds.Width;
            background.Height = bounds.Height;
            background.FillColor = "aliceBlue";

            baseGroup.Bounds.Width = bounds.Width;
            baseGroup.Bounds.Height = bounds.Height;

            var textBody = new SvgTextBodyRenderItem(baseGroup.Bounds, true);
            var paragraph = textBody.AddParagraph("Hello");

            paragraph.AddText(" There");

            var rtItem = new RichTextFormatSimple("Second paragraph", "Archivo Narrow", 16f, true);
            rtItem.FontColor = Color.DarkGreen;
            var para2 = textBody.AddParagraph(rtItem);

            baseGroup.AddChildItem(textBody);
            baseGroup.AddChildItem(background);

            List<RenderItem> items = new List<RenderItem>() { baseGroup };
            textBody.AppendRenderItems(items);

            svgShapeRenderer.Render(items);

            var svg = sb.ToString();


            SaveTextFileToWorkbook("svg\\textBodyStandAlone.svg", svg);
        }
    }
}
