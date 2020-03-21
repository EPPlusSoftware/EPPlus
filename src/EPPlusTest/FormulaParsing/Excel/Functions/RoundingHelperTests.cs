using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestClass]
    public class RoundingHelperTests
    {
        [TestMethod]
        public void CeilingShouldReturnCorrectResult()
        {
            var result = RoundingHelper.Round(22.25, 0.1, RoundingHelper.Direction.Up);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(22.25, 0.5, RoundingHelper.Direction.Up);
            Assert.AreEqual(22.5, result);

            result = RoundingHelper.Round(22.25, 1, RoundingHelper.Direction.Up);
            Assert.AreEqual(23, result);

            result = RoundingHelper.Round(22.25, 10, RoundingHelper.Direction.Up);
            Assert.AreEqual(30, result);

            result = RoundingHelper.Round(22.25, 20, RoundingHelper.Direction.Up);
            Assert.AreEqual(40, result);

            result = RoundingHelper.Round(-22.25, -0.1, RoundingHelper.Direction.Up);
            Assert.AreEqual(-22.3, result);

            result = RoundingHelper.Round(-22.25, -1, RoundingHelper.Direction.Up);
            Assert.AreEqual(-23, result);

            result = RoundingHelper.Round(-22.25, -5, RoundingHelper.Direction.Up);
            Assert.AreEqual(-25, result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResult_Down()
        {
            var result = RoundingHelper.Round(26.75, 0.1, RoundingHelper.Direction.Down);
            Assert.AreEqual(26.7, result);

            result = RoundingHelper.Round(26.75, 0.5, RoundingHelper.Direction.Down);
            Assert.AreEqual(26.5, result);

            result = RoundingHelper.Round(26.75, 1, RoundingHelper.Direction.Down);
            Assert.AreEqual(26, result);

            result = RoundingHelper.Round(26.75, 10, RoundingHelper.Direction.Down);
            Assert.AreEqual(20, result);

            result = RoundingHelper.Round(26.75, 20, RoundingHelper.Direction.Down);
            Assert.AreEqual(20, result);

            result = RoundingHelper.Round(-26.75, -0.1, RoundingHelper.Direction.Down);
            Assert.AreEqual(-26.7, result);

            result = RoundingHelper.Round(-26.75, -1, RoundingHelper.Direction.Down);
            Assert.AreEqual(-26, result);

            result = RoundingHelper.Round(-26.75, -5, RoundingHelper.Direction.Down);
            Assert.AreEqual(-25, result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResult_AlwaysDown()
        {
            var result = RoundingHelper.Round(26.75, 0.1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26.7, result);

            result = RoundingHelper.Round(26.75, 0.5, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26.5, result);

            result = RoundingHelper.Round(26.75, 1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26, result);

            result = RoundingHelper.Round(26.75, 10, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(20, result);

            result = RoundingHelper.Round(26.75, 0, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(0, result);

            result = RoundingHelper.Round(-26.25, -0.5, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(-26.5, result);

            result = RoundingHelper.Round(-26.75, 1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(-27, result);

            result = RoundingHelper.Round(-26.75, -1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(-27, result);

            result = RoundingHelper.Round(-26.75, 5, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(-30, result);
        }

        [TestMethod]
        public void CeilingShouldReturnCorrectResult_AlwaysUp()
        {
            var result = RoundingHelper.Round(22.25, 0.1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(22.25, 0.5, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(22.5, result);

            result = RoundingHelper.Round(22.25, -0.5, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(22.5, result);

            result = RoundingHelper.Round(22.25, 1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(23, result);

            result = RoundingHelper.Round(22.25, 10, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(30, result);

            result = RoundingHelper.Round(22.25, 0, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(0, result);

            result = RoundingHelper.Round(-22.25, -0.5, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(-22, result);

            result = RoundingHelper.Round(-22.25, 1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(-22, result);

            result = RoundingHelper.Round(-22.25, -1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(-22, result);

            result = RoundingHelper.Round(-22.25, 5, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(-20, result);
        }

        [TestMethod]
        public void NearestRoundingTest()
        {
            var result = RoundingHelper.Round(22.24, 0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(22.2, result);

            result = RoundingHelper.Round(22.25, 0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(22.26, 0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(-22.25, -0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(-22.3, result);

            result = RoundingHelper.Round(-22.24, -0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(-22.2, result);

            result = RoundingHelper.Round(333.8, 1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(334, result);

            result = RoundingHelper.Round(333.3, 1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(333, result);

            result = RoundingHelper.Round(333.3, 2, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(334, result);

            result = RoundingHelper.Round(555.3, 400, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(400, result);

            result = RoundingHelper.Round(555, 1000, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(1000, result);

            result = RoundingHelper.Round(-555.7, -1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(-556, result);

            result = RoundingHelper.Round(-555.4, -1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(-555, result);

            result = RoundingHelper.Round(-1555, -1000, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(-2000, result);
        }
    }
}
