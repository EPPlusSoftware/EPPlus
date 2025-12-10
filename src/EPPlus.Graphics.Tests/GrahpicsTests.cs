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
        public void SettingGlobalPositionOfParent()
        {
            Transform p1 = new Transform(new Vector2(3,3), Vector2.One);

            //This gets an all new vector2 with the position of p1 in the world
            var worldPosition = p1.Position;
            worldPosition.Y = 5;
            worldPosition.X = 10;

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

            p1.LocalPosition.Y = 5;
            Assert.AreEqual(p1.Position.Y, 5);
            Assert.AreEqual(5, p1.LocalPosition.Y);
        }
    }
}
