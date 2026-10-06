using EPPlus.Graphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using EPPlus.DrawingRenderer.RenderItems;

namespace EPPlus.DrawingRenderer.Tests.Shared
{
    [TestClass]
    public class MultiContainerTest
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

        [TestMethod]
        public void VerifyGrandChildContentTopLeftNegative()
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

            grandChild.Left = -5;
            grandChild.Top = -9;

            grandChild.Width = 15;
            grandChild.Height = 2;

            Assert.AreEqual(-2, container.ContentLeft);
            Assert.AreEqual(-6, container.ContentTop);
            Assert.AreEqual(30, container.ContentWidth);
            Assert.AreEqual(19 + 6, container.ContentHeight);
            Assert.AreEqual(28, container.ContentRight);
            Assert.AreEqual(19, container.ContentBottom);
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
    }
}
