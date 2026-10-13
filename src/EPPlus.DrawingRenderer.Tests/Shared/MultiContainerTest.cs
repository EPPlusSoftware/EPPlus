using EPPlus.DrawingRenderer.RenderItems;
using EPPlus.DrawingRenderer.Svg;
using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.DrawingRenderer.Tests.Shared
{
    [TestClass]
    public class MultiContainerTest : TestBase
    {
        [TestMethod]
        public void VerifySimpleChildContent()
        {
            var world = new BoundingBox(0, 0, 500, 500);

            var container = new MultiContainerItem(world);

            container.Left = 5;
            container.Top = 5;

            var child1 = new BoundingBox();
            child1.Left = 3;
            child1.Top = 3;

            child1.Width = 25;
            child1.Height = 16;

            child1.Parent = container;

            Assert.AreEqual(3, container.ContentLeft);
            Assert.AreEqual(3, container.ContentTop);
            Assert.AreEqual(child1.Width, container.ContentWidth);
            Assert.AreEqual(child1.Height, container.ContentHeight);
            Assert.AreEqual(child1.Width + child1.Left, container.ContentRight);
            Assert.AreEqual(child1.Height + child1.Top, container.ContentBottom);
        }

        [TestMethod]
        public void VerifyGrandChildContentWidthHeight()
        {
            var world = new BoundingBox(0, 0, 500, 500);

            var container = new MultiContainerItem(world);

            container.Left = 5;
            container.Top = 5;

            var child1 = new BoundingBox();
            child1.Left = 3;
            child1.Top = 3;

            child1.Width = 25;
            child1.Height = 16;

            child1.Parent = container;

            var grandChild = new BoundingBox();
            grandChild.Parent = child1;

            grandChild.Width = 60;
            grandChild.Height = 100;

            Assert.AreEqual(3, container.ContentLeft);
            Assert.AreEqual(3, container.ContentTop);
            Assert.AreEqual(60, container.ContentWidth);
            Assert.AreEqual(100, container.ContentHeight);
            Assert.AreEqual(63, container.ContentRight);
            Assert.AreEqual(103, container.ContentBottom);
        }

        private void GenerateSvgFile(string fileName, BoundingBox bounds, params Transform[] items)
        {

            StringBuilder sb = new StringBuilder();
            var svgShapeRenderer = new SvgShapeRenderer(bounds, sb, new SvgRenderOptions());

            List<Transform> renderItems = items.ToList();
            svgShapeRenderer.Render(renderItems);

            var svg = sb.ToString();

            SaveTextFileToWorkbook($"svg\\{fileName}.svg", svg);
        }

        [TestMethod]
        public void VerifySimpleChildContentNegativeTopLeft()
        {
            var world = new BoundingBox(0, 0, 500, 500);

            var container = new MultiContainerItem(world);

            container.Left = 5;
            container.Top = 5;

            var child1 = new BoundingBox();
            child1.Left = -3;
            child1.Top = -3;

            child1.Width = 25;
            child1.Height = 16;

            child1.Parent = container;

            Assert.AreEqual(child1.Left, container.ContentLeft);
            Assert.AreEqual(child1.Top, container.ContentTop);
            Assert.AreEqual(child1.Width, container.ContentWidth);
            Assert.AreEqual(child1.Height, container.ContentHeight);
            Assert.AreEqual(child1.Width + child1.Left, container.ContentRight);
            Assert.AreEqual(child1.Height + child1.Top, container.ContentBottom);
        }

        [TestMethod]
        public void VerifyGrandChildContentTopLeftNegative()
        {
            var world = new GroupRenderItem();
            world.Width = 500;
            world.Height = 500;

            var container = new MultiContainerItem(world);

            container.Left = 5;
            container.Top = 5;

            var child1 = new GroupRenderItem();
            child1.Left = 3;
            child1.Top = 3;

            child1.Width = 25;
            child1.Height = 16;

            child1.Parent = container;

            var grandChild = new GroupRenderItem();
            grandChild.Parent = child1;

            grandChild.Left = -5;
            //This puts grandchild Y at -1 globally
            //A bit weird but technically allowed in e.g. SVG
            grandChild.Top = -9;

            grandChild.Width = 15;
            grandChild.Height = 2;

            Assert.AreEqual(-2, container.ContentLeft);
            Assert.AreEqual(-6, container.ContentTop);
            Assert.AreEqual(30, container.ContentWidth);
            //In this Case.
            //and the distance between 3(child) + 16 (child height)
            //The distance between -9 (grandChild) and 3(child) is 6
            //Alternatively. Globally our extremes are at -1 and 24 giving a total distance of ABS(-1-24) = 25
            Assert.AreEqual(19 + 6, container.ContentHeight);
            Assert.AreEqual(28, container.ContentRight);
            Assert.AreEqual(19, container.ContentBottom);

            //As this may be hard to visualize here's a very simple svg file that illustrates this setup:
            /*
                <svg class="world" xmlns="http://www.w3.org/2000/svg" width="500" height="500" viewBox="-1 -1 500 500" fill="none">
                  <g class="container" transform="translate(5,5)">
                    <g class="child1" transform="translate(3,3)">
                      <rect class="child1_visualization" width="25" height="16" fill="red" />
                      <g class="child2" transform ="translate(-5,-9)">
                        <rect class="child2_visualization" width="15" height="2" fill="yellow" />
                      </g>
                    </g>
                  </g>
                  <rect class="TotalSize_GlobalRect" fill="green" x="3" y="-1" fill-opacity="20%" width="30" height="25"></rect>
                  <rect class="ExtremeRightBottom_GlobalRect" fill="blue" x="33" y="24" width="5" height="5"></rect>
                </svg>
             */

            //If you wish you can also generate a similar file from the existing structure:
            //RectRenderItem totalSize = new RectRenderItem(world);
            //totalSize.Name = "TotalSize_GlobalRect";
            //totalSize.Left = 3;
            //totalSize.Top = -1;
            //totalSize.Width = 30;
            //totalSize.Height = 25;
            //totalSize.Style.FillColor = "green";
            //totalSize.Style.FillOpacity = 0.2d;

            ////Will be drawing a rect which has it's TopLeft in the RightBottom of the case above
            //RectRenderItem ExtremeRightBottom = new RectRenderItem(world);
            //ExtremeRightBottom.Name = "ExtremeRightBottom_GlobalRect";
            //ExtremeRightBottom.Left = 33;
            //ExtremeRightBottom.Top = 24;
            //ExtremeRightBottom.Width = 5;
            //ExtremeRightBottom.Height = 5;
            //ExtremeRightBottom.Style.FillColor = "blue";
            //ExtremeRightBottom.Style.FillOpacity = 0.2d;

            //GenerateSvgFile("VerifyGrandChildContentTopLeftNegative", world, world.ChildObjects.ToArray());
        }
    }
}
