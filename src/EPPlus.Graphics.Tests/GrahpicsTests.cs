using EPPlus.Graphics;
using EPPlus.Graphics.Units;
using EPPlus.Graphics.Math;
using System.Drawing;

namespace EPPlus.Graphics.Tests
{
    [TestClass]
    public class GraphicsTests
    {
        [TestMethod]
        public void AddChildTest()
        {
            Transform p1 = new Transform();
            Transform c1 = new Transform(Vector2.One, Vector2.One, p1);
            Assert.HasCount(1, p1.ChildObjects);
        }

        [TestMethod]
        public void MoveParentTransformTest()
        {
            Transform p1 = new Transform();
            Transform c1 = new Transform(Vector2.One, Vector2.One, p1);
            p1.Translate(Vector2.One);
            Assert.AreEqual(2, c1.Position.X);
            Assert.AreEqual(1, c1.LocalPosition.X);
        }

        [TestMethod]
        public void ScaleParentTransformTest()
        {
            Transform p1 = new Transform();
            Transform c1 = new Transform(Vector2.One, Vector2.One, p1);
            p1.Scale = new Vector2(2,2);
            Assert.AreEqual(1, c1.LocalScale.X);
            Assert.AreEqual(2, c1.Scale.X);
        }

        [TestMethod]
        public void ChildAndGrandChildTest()
        {
            Transform p1 = new Transform();
            Transform c1 = new Transform(Vector2.One, Vector2.One, p1);
            var grandChildPosition = new Vector2(1.5, 2.33333333);

            Transform c2 = new Transform(grandChildPosition, Vector2.One, c1);

            var posOffset = new Vector2(0.2, 1.1);
            p1.Translate(posOffset);

            Assert.AreEqual(2.7, c2.Position.X);
            Assert.AreEqual(4.43333333, c2.Position.Y);
            Assert.AreEqual(1.5, c2.LocalPosition.X);
            Assert.AreEqual(2.33333333, c2.LocalPosition.Y);
        }

        [TestMethod]
        public void ChildAndGrandChildTestLocalPos()
        {
            Transform p1 = new Transform();
            p1.LocalPosition = new Vector2(2, 2);

            Transform c1 = new Transform();
            c1.Parent = p1;

            c1.LocalPosition = new Vector2(5, 5);

            Transform gc1 = new Transform();

            //Statics in vector2 are unintentionally remaining?
            Assert.AreEqual(0, gc1.LocalPosition.X);
            Assert.AreEqual(0, gc1.LocalPosition.Y);
        }

        [TestMethod]
        public void ChildAndGrandChildTestLocalPosOtherConstructor()
        {
            var startPos = new Vector2(2, 2);
            Transform p1 = new Transform(startPos, Vector2.One);

            var localStartPosC1 = new Vector2(5, 5);
            Transform c1 = new Transform();
            c1.Parent = p1;

            c1.LocalPosition = localStartPosC1;

            Transform gc1 = new Transform();

            gc1.Parent = c1;

            var localGC1 = new Vector2(10, 10);
            gc1.LocalPosition = localGC1;

            Assert.AreEqual(2, p1.LocalPosition.X);
            Assert.AreEqual(2, p1.LocalPosition.Y);

            Assert.AreEqual(5, c1.LocalPosition.X);
            Assert.AreEqual(5, c1.LocalPosition.Y);

            Assert.AreEqual(10, gc1.LocalPosition.X);
            Assert.AreEqual(10, gc1.LocalPosition.Y);
        }

        [TestMethod]
        public void SettingGlobalPositionOfParent()
        {
            Transform p1 = new Transform(new Vector2(3,3), Vector2.One);

            //This gets an all new vector2 with the position of p1 in the world
            var worldPosition = p1.Position;
            worldPosition.Y = 5;
            worldPosition.X = 10;

            p1.Position = worldPosition;

            Assert.AreEqual(5, p1.LocalPosition.Y);
            Assert.AreEqual(5, p1.Position.Y);
        }


        [TestMethod]
        public void SettingLocalPositionOfBase()
        {
            Transform p1 = new Transform(new Vector2(3, 3), Vector2.One);

            Assert.AreEqual(3, p1.Position.Y);
            Assert.AreEqual(3, p1.LocalPosition.Y);
            Assert.AreEqual(p1.Position.Y, p1.LocalPosition.Y);

            p1.LocalPosition = new(p1.LocalPosition.X, 5);
            Assert.AreEqual(p1.Position.Y, 5);
            Assert.AreEqual(5, p1.LocalPosition.Y);
        }

        [TestMethod]
        public void BoundingBoxes()
        {
            BoundingBox Shape = new BoundingBox();

            Shape.Position = new Vector2(2,2);
            Shape.Size = Vector2.One;
            Shape.Name = "ShapeTransform";

            Shape.Width = 10;
            Shape.Height = 10;

            BoundingBox TextBody = new BoundingBox();

            TextBody.Parent = Shape;
            TextBody.Name = "TextBodyTransform";

            TextBody.Width = 20;
            TextBody.Height = 20;

            TextBody.LocalPosition = new(10, 11);

            Assert.AreEqual(10, TextBody.LocalPosition.X);
            Assert.AreEqual(11, TextBody.LocalPosition.Y);

            BoundingBox Paragraph1 = new BoundingBox();
            Paragraph1.Name = "Paragraph1Transform";
            Paragraph1.Parent = TextBody;
            Paragraph1.Width = 5;
            Paragraph1.Height = 5;


            Paragraph1.LocalPosition = new Vector2(5, 8);

            Assert.AreEqual(10, TextBody.LocalPosition.X);
            Assert.AreEqual(11, TextBody.LocalPosition.Y);

            Assert.AreEqual(2, Shape.LocalPosition.X);
            Assert.AreEqual(2, Shape.LocalPosition.Y);

            Assert.AreEqual(10, Shape.ChildObjects[0].LocalPosition.X);
            Assert.AreEqual(11, Shape.ChildObjects[0].LocalPosition.Y);

            Assert.AreEqual(5, Paragraph1.LocalPosition.X);
            Assert.AreEqual(8, Paragraph1.LocalPosition.Y);

            Assert.AreEqual(Paragraph1.Position.X, 17);
            Assert.AreEqual(Paragraph1.Position.Y, 21);


            var str = Shape.ToHierarchyString();
        }
    }
}
