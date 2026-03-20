using EPPlus.Export.ImageRenderer.RenderItems.SvgItem;
using EPPlus.Export.ImageRenderer.Svg;
using EPPlus.Graphics;
using EPPlusImageRenderer.RenderItems;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Tests
{
    [TestClass]
    public class GroupItemNewTest : TestBase
    {
        [TestMethod]
        public void TestGroupItemMovingTwoChildrenCorrectly()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var baseItem = new DrawingItemForTesting(baseBB);

            var groupItem = new SvgGroupItemNew(baseItem, 9, 9);

            SvgRenderRectItem rectItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);

            rectItem.FillColor = "red";
            rectItem.FillOpacity = 0.2d;

            rectItem.Width = 20;
            rectItem.Height = 20;

            groupItem.AddChildItem(rectItem);


            SvgRenderRectItem siblingItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);
            siblingItem.FillColor = "blue";
            siblingItem.FillOpacity = 0.2d;

            siblingItem.Width = 20;
            siblingItem.Height = 20;

            siblingItem.Bounds.Left = 20;
            siblingItem.Bounds.Top = 20;

            groupItem.AddChildItem(siblingItem);

            var worldCoordinatesRectBefore = rectItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(9, worldCoordinatesRectBefore.X);
            Assert.AreEqual(9, worldCoordinatesRectBefore.Y);

            var worldCoordinatesSibBefore = siblingItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(29, worldCoordinatesSibBefore.X);
            Assert.AreEqual(29, worldCoordinatesSibBefore.Y);

            groupItem.Position.Left = 18;
            groupItem.Position.Top = 18;

            var worldCoordinatesRect = rectItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(18, worldCoordinatesRect.X);
            Assert.AreEqual(18, worldCoordinatesRect.Y);

            var worldCoordinatesSib = siblingItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(38, worldCoordinatesSib.X);
            Assert.AreEqual(38, worldCoordinatesSib.Y);

            var sb = new StringBuilder();

            baseItem.RenderItems.Add(groupItem);

            baseItem.Render(sb);
            var svgString = sb.ToString();

            SaveTextFileToWorkbook($"svg\\StandAloneTestGroup.svg", svgString);
        }

        [TestMethod]
        public void TestGroupItemRotatingTwoChildrenCorrectly()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var baseItem = new DrawingItemForTesting(baseBB);

            BoundingBox parent = new BoundingBox();

            var groupItem = new SvgGroupItemNew(baseItem, parent, 45);

            groupItem.Position.Left = 10;
            groupItem.Position.Top = 10;

            SvgRenderRectItem rectItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);

            rectItem.FillColor = "red";
            rectItem.FillOpacity = 0.2d;

            rectItem.Width = 20;
            rectItem.Height = 20;

            groupItem.AddChildItem(rectItem);


            SvgRenderRectItem siblingItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);
            siblingItem.FillColor = "blue";
            siblingItem.FillOpacity = 0.2d;

            siblingItem.Width = 20;
            siblingItem.Height = 20;

            siblingItem.Bounds.Left = 20;
            siblingItem.Bounds.Top = 20;

            groupItem.AddChildItem(siblingItem);

            parent.Width = groupItem.Bounds.Width;
            parent.Height = groupItem.Bounds.Height;

            var sb = new StringBuilder();

            baseItem.RenderItems.Add(groupItem);

            baseItem.Render(sb);
            var svgString = sb.ToString();

            SaveTextFileToWorkbook($"svg\\StandAloneTestGroup.svg", svgString);
        }
    }
}
