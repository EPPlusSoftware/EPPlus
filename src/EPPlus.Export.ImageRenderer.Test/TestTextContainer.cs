using EPPlus.Export.ImageRenderer.Text;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Interfaces.Drawing.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EPPlus.Graphics;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType;
using EPPlus.Graphics;
using EPPlus.Export.ImageRenderer.RenderItems.Shared;
using EPPlus.Fonts.OpenType;

namespace EPPlus.Export.ImageRenderer.Tests
{
    [TestClass]
    public class TestTextContainer
    {
        [TestMethod]
        public void TestDynamicSizeSingleLine()
        {
            string content = "TextBox";
            var expectedWidth = 48d;
            var expectedHeight = 18d;

            var mf = new MeasurementFont() { FontFamily = "Aptos", Size = 11, Style = MeasurementFontStyles.Regular };
            
            var container = new TextContainer(content, mf, true);

            Assert.AreEqual(expectedWidth, Math.Round(container.Width,0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height,0));
        }

        [TestMethod]
        public void TestDynamicSizeMultipleLines()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            var expectedWidth = 93d;
            var expectedHeight = 54d;

            //0.57 inches
            //0.98 inches

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 11, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true);

            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void TestNonStandardFontSizesMultiLine()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            var expectedWidth = 203d;
            var expectedHeight = 117d;

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 24, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true, true);
            //0,056640625  * font size width correction for pixels (minimum)
            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void TestNonStandardFontSizesMultiLineLargeFont()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            var expectedWidth = 813d;
            var expectedHeight = 469d;

            var mf = new MeasurementFont() { FontFamily = "Aptos Narrow", Size = 96, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true, true);
            //0,072265625 * font size width correction for pixels
            //0,0442708333333333 * font size height correction for pixels
            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void TestNonStandardFontSizesMultiLineLargeFontGoudyStout()
        {
            string content = "TextBox\r\na very long line2\r\nline3";
            //This is the actual height in excel.
            //var expectedHeight = 533d;
            //It is unclear why as the Max possible glyph height in pixels for each line is
            //175.125 pixels. *3 = 525.375
            //Best guess is it adds 4 pixels per row for potential "outline"

            //This is the glyph yMax and YMin
            var expectedHeight = 525d;
            //Excel has 2332d but they add some kind of internal buffer/minX spacing for lineEnding
            var expectedWidth = 2309d;

            var mf = new MeasurementFont() { FontFamily = "Goudy Stout", Size = 96, Style = MeasurementFontStyles.Regular };

            var container = new TextContainer(content, mf, true, true);
            //0,072265625 * font size width correction for pixels
            //0,0442708333333333 * font size height correction for pixels
            Assert.AreEqual(expectedWidth, Math.Round(container.Width, 0));
            Assert.AreEqual(expectedHeight, Math.Round(container.Height, 0));
        }

        [TestMethod]
        public void EnsureTextBodyAddsRunsCorrectly()
        {
            BoundingBox shapeRect = new BoundingBox();

            shapeRect.Width = 20;
            shapeRect.Height = 10;

            FontMeasurerTrueType measurer = new FontMeasurerTrueType(12, "Aptos Narrow", FontSubFamily.Regular);
            var body = new TextBody(shapeRect);

            body.Bounds.transform.Name = "TxtBody";

            body.AddText("A new Paragraph", measurer);
            body.AddText("Second paragraph", measurer);

            Assert.AreEqual(2, body.Bounds.transform.ChildObjects.Count);
            Assert.AreEqual(body.Bounds.transform.ChildObjects[0], body.Paragraphs[0].Bounds.transform);

            Assert.AreEqual(shapeRect.transform, body.Bounds.transform.Parent);
        }

        //[TestMethod]
        //public void TextContainerGeneric()
        //{
        //    //Options

        //    //1: One text-container per "shape/textbox"
        //    //Pros: Resizes based on strings in a singular place, less "New" statements, Less broken down
        //    //Cons: Individual fonts etc. get tricky, will still need to be broken down but more obfuscated, 

        //    //2: Text-Container down to fragment/run level
        //    //Pros: All positioning, fonts, etc. Come in by default. Each fragment knows where it is and what it is and how big it is, need only ever measure each fragment once
        //    //Cons: Text-wrapping, resizing the shape, etc gets trickier, More new statements means larger overhead and less clear overview

        //    //3: Text-container down to Paragraph level
        //    //Pros: All paragraph properties; indentation, linespacing, alignment etc gets well contained and clear, 
        //    //Cons: 

        //    //I'm thinking of this all wrong.
        //    //Either way we will need a transform for each object in a hierarchy.
        //    //Lowest common denominator has to become some kind of baseclass
        //    //Build from smallest (fragment) to Largest. All of the same base-class (Could be as simple as just containers for Transform for starters
        //    //Then go upward. (Glyph?) -> Fragment/Run -> Run -> Paragraph -> Container
        //    //Lowest needs: Rect. (Bounding box) that can also hold a string.
        //    //Next: Same but basic font data
        //    //Next: Full rich-text?

        //    StringBuilder sb = new StringBuilder();

        //    TextContainerBase Fragment1 = new TextContainerBase(true);

        //    TextContainerBase Fragment2 = new TextContainerBase();

        //    TextContainerBase Fragment3 = new TextContainerBase();
        //    ////Test positioning and re-sizing
        //    //TextContainer shapeTextBox = new TextContainer();

        //    //shapeTextBox.transform.Position = new Graphics.Math.Vector2(5, 10);


        //    //TextContainer paragraph1Container = new TextContainer();

        //    //TextContainer para2 = new TextContainer();

        //    //TextContainer para3 = new TextContainer();

        //    //TextContainer para4 = new TextContainer();
        //}
    }
}
