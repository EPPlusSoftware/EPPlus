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

            var textBody = new RenderTextBody(baseGroup.Bounds, true);

            textBody.Text = "Hello";
            //var para = new SvgParagraphRenderItem(textBody, textBody.Bounds);
            
            //var para2 = new DrawingParagraphRenderItem(textBody, textBody.Bounds);
            //textBody.Paragraphs.Add

            baseGroup.AddChildItem(rectItem);
            baseGroup.AddChildItem(textBody);

            List<RenderItem> items = new List<RenderItem>() { baseGroup };

            svgShapeRenderer.Render(items);

            var svg = sb.ToString();

            var fi = GetOutputFile("svg/", "RectTestStandAlone.svg");
            SaveTextFileToWorkbook("svg\\rectStandalone.svg", svg);
        }

        [TestMethod]
        public void SvgTextBoxTest()
        {
            BoundingBox bounds = new BoundingBox(0, 0, 500, 500);
            StringBuilder sb = new StringBuilder();
            var svgShapeRenderer = new SvgShapeRenderer(bounds, sb);
            

            var baseGroup = new GroupRenderItem(bounds);

            var background = new RectRenderItem(baseGroup.Bounds);

            background.Width = bounds.Width;
            background.Height = bounds.Height;
            background.FillColor = "aliceBlue";

            var textBody = new RenderTextBody(baseGroup.Bounds, true);


            textBody.Text = "Hello";
            var para = new SvgParagraphRenderItem(textBody, textBody.Bounds);

           
            //var para2 = new DrawingParagraphRenderItem(textBody, textBody.Bounds);
            //textBody.Paragraphs.Add


            //para.Runs.Add()

            baseGroup.AddChildItem(background);
            baseGroup.AddChildItem(textBody);

            List<RenderItem> items = new List<RenderItem>() { baseGroup };

            svgShapeRenderer.Render(items);

            var svg = sb.ToString();

            var fi = GetOutputFile("svg/", "RectTestStandAlone.svg");
            SaveTextFileToWorkbook("svg\\rectStandalone.svg", svg);
        }
    }
}
