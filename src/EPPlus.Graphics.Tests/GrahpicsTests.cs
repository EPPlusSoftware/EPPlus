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
    }
}
