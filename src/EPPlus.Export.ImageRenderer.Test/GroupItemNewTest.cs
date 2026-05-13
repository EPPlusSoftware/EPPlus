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
        //private void InitalizeTransformGroup()
        //{
        //    var baseBB = new BoundingBox();

        //    //96x96 px
        //    baseBB.Width = 72;
        //    baseBB.Height = 72;

        //    var baseItem = new DrawingItemForTesting(baseBB);

        //    var groupItem = new SvgTransformGroup(baseItem, 9, 9);

        //    SvgRenderRectItem rectItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);

        //    rectItem.FillColor = "red";
        //    rectItem.FillOpacity = 0.2d;

        //    rectItem.Width = 20;
        //    rectItem.Height = 20;

        //    groupItem.AddChildItem(rectItem);


        //    SvgRenderRectItem siblingItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);
        //    siblingItem.FillColor = "blue";
        //    siblingItem.FillOpacity = 0.2d;

        //    siblingItem.Width = 20;
        //    siblingItem.Height = 20;

        //    siblingItem.Bounds.Left = 20;
        //    siblingItem.Bounds.Top = 20;

        //    groupItem.AddChildItem(siblingItem);
        //}


        [TestMethod]
        public void TransformGroupMoveCorrect()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var baseItem = new DrawingItemForTesting(baseBB);

            var holderBB = new BoundingBox();
            holderBB.Parent = baseItem.Bounds;

            var groupItem = new SvgTransformGroup(baseItem, 9, 9);
            groupItem.Bounds.Parent = holderBB;

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

            groupItem.Bounds.Left = 18;
            groupItem.Bounds.Top = 18;

            var worldCoordinatesRect = rectItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(18, worldCoordinatesRect.X);
            Assert.AreEqual(18, worldCoordinatesRect.Y);

            var worldCoordinatesSib = siblingItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(38, worldCoordinatesSib.X);
            Assert.AreEqual(38, worldCoordinatesSib.Y);

            groupItem.Scale = new EPPlusImageRenderer.Coordinate(0.5d, 0.5d);

            groupItem.Bounds.Left = 0;
            groupItem.Bounds.Top = 0;

            //groupItem.Bounds.Left = 0;
            //groupItem.Bounds.Top = 0;

            rectItem.Bounds.Left = 18;
            rectItem.Bounds.Top = 18;

            var unchangedVector = rectItem.Bounds.GetWorldCoordinates();

            ////Graphics.Math.Matrix3x3 scaleHalfMatrix
            var worldCoordinatesRectAfterScale = rectItem.Bounds.Position;
            Assert.AreEqual(9, worldCoordinatesRectAfterScale.X);
            Assert.AreEqual(9, worldCoordinatesRectAfterScale.Y);

            var worldCoordinatesSibAfterScale = siblingItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(29, worldCoordinatesSibAfterScale.X);
            Assert.AreEqual(19, worldCoordinatesSibAfterScale.Y);


            var sb = new StringBuilder();

            baseItem.RenderItems.Add(groupItem);

            baseItem.Render(sb);
            var svgString = sb.ToString();

            SaveTextFileToWorkbook($"svg\\StandAloneTranslateGroup.svg", svgString);
        }


        [TestMethod]
        public void GroupInGroupMoveCorrect2()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var baseItem = new DrawingItemForTesting(baseBB);

            var groupItem = new SvgGroupItemNew(baseItem, 5, 15);

            SvgRenderRectItem rectItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);

            groupItem.AddChildItem(rectItem);

            rectItem.FillColor = "red";
            rectItem.FillOpacity = 0.2d;

            rectItem.Width = 20;
            rectItem.Height = 20;

            groupItem.TranslationOffset.Left = 5;
            groupItem.TranslationOffset.Top = 6;

            var leftGlobal = groupItem.TranslationOffset.Position.X;
            var topGlobal = groupItem.TranslationOffset.Position.Y;
            var leftGlobalUnder = groupItem.Bounds.Position.X;
            var topGlobalUnder = groupItem.Bounds.Position.Y;

            Assert.AreEqual(10, leftGlobalUnder);
            Assert.AreEqual(21, topGlobalUnder);
        }

        [TestMethod]
        public void GroupInGroupMoveCorrect()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var baseItem = new DrawingItemForTesting(baseBB);

            var groupItem = new SvgGroupItemNew(baseItem, 9, 9);

            var subItem = new SvgGroupItemNew(baseItem, 9, 9);
            subItem.TranslationOffset.Parent = groupItem.TranslationOffset;
            groupItem.AddChildItem(subItem);

            SvgRenderRectItem rectItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);

            rectItem.FillColor = "red";
            rectItem.FillOpacity = 0.2d;

            rectItem.Width = 20;
            rectItem.Height = 20;

            subItem.AddChildItem(rectItem);

            SvgRenderRectItem siblingItem = new SvgRenderRectItem(baseItem, groupItem.Bounds);
            siblingItem.FillColor = "blue";
            siblingItem.FillOpacity = 0.2d;

            siblingItem.Width = 20;
            siblingItem.Height = 20;

            siblingItem.Bounds.Left = 20;
            siblingItem.Bounds.Top = 20;

            subItem.AddChildItem(siblingItem);

            var worldCoordinatesRectBefore = rectItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(18, worldCoordinatesRectBefore.X);
            Assert.AreEqual(18, worldCoordinatesRectBefore.Y);

            var worldCoordinatesSibBefore = siblingItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(38, worldCoordinatesSibBefore.X);
            Assert.AreEqual(38, worldCoordinatesSibBefore.Y);

            subItem.TranslationOffset.Left = 9;
            subItem.TranslationOffset.Top = 9;

            var worldCoordinatesRect = rectItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(27, worldCoordinatesRect.X);
            Assert.AreEqual(27, worldCoordinatesRect.Y);

            var worldCoordinatesSib = siblingItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(47, worldCoordinatesSib.X);
            Assert.AreEqual(47, worldCoordinatesSib.Y);

            var sb = new StringBuilder();

            baseItem.RenderItems.Add(groupItem);

            baseItem.Render(sb);
            var svgString = sb.ToString();

            SaveTextFileToWorkbook($"svg\\subItemInSubItem.svg", svgString);
        }

        [TestMethod]
        public void GrpPosTranslateChildren()
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

            groupItem.TranslationOffset.Left = 9;
            groupItem.TranslationOffset.Top = 9;

            var worldCoordinatesRect = rectItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(18, worldCoordinatesRect.X);
            Assert.AreEqual(18, worldCoordinatesRect.Y);

            var worldCoordinatesSib = siblingItem.Bounds.GetWorldCoordinates();
            Assert.AreEqual(38, worldCoordinatesSib.X);
            Assert.AreEqual(38, worldCoordinatesSib.Y);

            var gLeft = groupItem.Bounds.Position.X;
            var gTop = groupItem.Bounds.Position.Y;

            var sb = new StringBuilder();

            baseItem.RenderItems.Add(groupItem);

            baseItem.Render(sb);
            var svgString = sb.ToString();

            SaveTextFileToWorkbook($"svg\\StandAloneTestGroup.svg", svgString);
        }

        [TestMethod]
        public void RotateTwoChildren()
        {
            var baseBB = new BoundingBox();

            //96x96 px
            baseBB.Width = 72;
            baseBB.Height = 72;

            var baseItem = new DrawingItemForTesting(baseBB);

            BoundingBox parent = new BoundingBox();

            var groupItem = new SvgGroupItemNew(baseItem, parent, 45);

            groupItem.TranslationOffset.Left = 10;
            groupItem.TranslationOffset.Top = 10;

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

            groupItem.SetRotationPointToCenterOfGroup();

            SvgRenderRectItem centerOfGroupMarker = new SvgRenderRectItem(baseItem, baseItem.Bounds);
            centerOfGroupMarker.FillColor = "green";
            centerOfGroupMarker.FillOpacity = 0.8d;

            centerOfGroupMarker.Width = 6;
            centerOfGroupMarker.Height = 6;

            centerOfGroupMarker.Left = 30 - (centerOfGroupMarker.Width / 2);
            centerOfGroupMarker.Top = 30 - (centerOfGroupMarker.Height / 2);

            baseItem.RenderItems.Add(centerOfGroupMarker);

            var sb = new StringBuilder();

            baseItem.RenderItems.Add(groupItem);

            baseItem.Render(sb);
            var svgString = sb.ToString();

            SaveTextFileToWorkbook($"svg\\TestGroupRotated.svg", svgString);
        }
    }
}
