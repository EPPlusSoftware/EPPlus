using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.ImageRenderer.Tests
{
    [TestClass]
    public class TextboxTest : TestBase
    {
        [TestMethod]
        public void TextBoxVerification()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var item = new DrawingItemForTesting(baseBB);

            BoundingBox maxBounds = new BoundingBox();
            maxBounds.Width = 36;
            maxBounds.Height = 36;

            var txtBox = new SvgTextBox(item, item.Bounds, maxBounds);
            txtBox.AddText(0, "My new text which is fun");
            txtBox.Rectangle.FillColor = "red";
            txtBox.Rectangle.FillOpacity = 0.2d;

            txtBox.Left = 5;
            txtBox.Top = 5;

            item.ExternalRenderItemsNoBounds.Add(txtBox);

            var txtBoxNotMaxed = new SvgTextBox(item, item.Bounds, maxBounds);
            txtBoxNotMaxed.Left = 42;
            txtBoxNotMaxed.Top = 5;

            txtBoxNotMaxed.AddText(0, "abc");
            txtBoxNotMaxed.Rectangle.FillColor = "yellow";
            txtBoxNotMaxed.Rectangle.FillOpacity = 0.2d;

            item.ExternalRenderItemsNoBounds.Add(txtBoxNotMaxed);

            var sb = new StringBuilder();

            item.Render(sb);
            var svgString = sb.ToString();

            //before we assumed we consider the space widths
            var widthWithSpace = txtBox.TextBody.Paragraphs[0].SpaceWidthsPerLine[0] + txtBox.Width;


            //Assert.AreEqual(36d, txtBox.TextBody.Width);
            Assert.AreEqual(37.5d, widthWithSpace, 0.5);
            Assert.AreNotEqual(36d, txtBoxNotMaxed.Width);
            Assert.AreEqual(16.29052734375d, txtBoxNotMaxed.Width);

            //var sb = new StringBuilder();

            //item.Render(sb);
            //var svgString = sb.ToString();

            SaveTextFileToWorkbook($"svg\\StandAloneTextBox.svg", svgString);
        }
    }
}
