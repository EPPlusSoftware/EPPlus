using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
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
            // Direction.Up = Ceiling toward +∞
            var result = RoundingHelper.Round(22.25, 0.1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(22.25, 0.5, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(22.5, result);

            result = RoundingHelper.Round(22.25, 1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(23, result);

            result = RoundingHelper.Round(22.25, 10, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(30, result);

            result = RoundingHelper.Round(22.25, 20, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(40, result);

            // Negatives: ceiling gives LESS negative (toward +∞)
            result = RoundingHelper.Round(-22.25, -0.1, RoundingHelper.Direction.AlwaysUp); // |multiple| = 0.1
            Assert.AreEqual(-22.3, result);

            result = RoundingHelper.Round(-22.25, -1, RoundingHelper.Direction.AlwaysUp);   // |multiple| = 1
            Assert.AreEqual(-23, result);

            result = RoundingHelper.Round(-22.25, -5, RoundingHelper.Direction.AlwaysUp);   // |multiple| = 5
            Assert.AreEqual(-20, result);

            // Edges
            result = RoundingHelper.Round(555, 1000, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(1000, result);

            result = RoundingHelper.Round(-555, -1000, RoundingHelper.Direction.AlwaysUp);  // |multiple| = 1000
            Assert.AreEqual(-1000, result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResult_Down()
        {
            // Direction.Down = Floor toward −∞
            var result = RoundingHelper.Round(26.75, 0.1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26.7, result);

            result = RoundingHelper.Round(26.75, 0.5, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26.5, result);

            result = RoundingHelper.Round(26.75, 1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26, result);

            result = RoundingHelper.Round(26.75, 10, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(20, result);

            result = RoundingHelper.Round(26.75, 20, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(20, result);

            // Negatives: floor gives MORE negative (away from +∞, toward −∞)
            result = RoundingHelper.Round(-26.75, -0.1, RoundingHelper.Direction.AlwaysDown); // |multiple| = 0.1
            Assert.AreEqual(-26.7, result);

            result = RoundingHelper.Round(-26.75, -1, RoundingHelper.Direction.AlwaysDown);   // |multiple| = 1
            Assert.AreEqual(-26, result);

            result = RoundingHelper.Round(-26.75, -5, RoundingHelper.Direction.AlwaysDown);   // |multiple| = 5
            Assert.AreEqual(-25, result);

            // Edges
            result = RoundingHelper.Round(555, 1000, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(0, result);

            result = RoundingHelper.Round(-555, -1000, RoundingHelper.Direction.AlwaysDown);  // |multiple| = 1000
            Assert.AreEqual(0, result);
        }

        [TestMethod]
        public void FloorShouldReturnCorrectResult_AlwaysDown()
        {
            // Direction.AlwaysDown = Toward zero (positives -> floor, negatives -> ceiling)
            const double eps = 1e-12;

            var result = RoundingHelper.Round(26.75, 0.1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26.7, result, eps);

            result = RoundingHelper.Round(26.75, 0.5, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26.5, result, eps);

            result = RoundingHelper.Round(26.75, 1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(26, result, eps);

            result = RoundingHelper.Round(26.75, 10, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(20, result, eps);

            // multiple == 0 -> 0 per implementation
            result = RoundingHelper.Round(26.75, 0, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(0, result, eps);

            // Negatives: toward zero => ceiling on quotient
            result = RoundingHelper.Round(-26.25, -0.5, RoundingHelper.Direction.AlwaysDown); // |multiple| = 0.5
            Assert.AreEqual(-26.0, result, eps);

            result = RoundingHelper.Round(-26.75, 1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(-27, result, eps);

            result = RoundingHelper.Round(-26.75, -1, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(-26, result, eps);

            result = RoundingHelper.Round(-26.75, 5, RoundingHelper.Direction.AlwaysDown);
            Assert.AreEqual(-25, result, eps);
        }

        [TestMethod]
        public void CeilingShouldReturnCorrectResult_AlwaysUp()
        {
            // Direction.AlwaysUp = Away from zero (positives -> ceiling, negatives -> floor)
            var result = RoundingHelper.Round(22.25, 0.1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(22.25, 0.5, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(22.5, result);

            result = RoundingHelper.Round(22.25, 1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(23, result);

            result = RoundingHelper.Round(22.25, 10, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(30, result);

            // multiple == 0 -> 0 per implementation
            result = RoundingHelper.Round(22.25, 0, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(0, result);

            // Negatives: away from zero => floor on quotient
            result = RoundingHelper.Round(-22.25, -0.5, RoundingHelper.Direction.AlwaysUp); // |multiple| = 0.5
            Assert.AreEqual(-22.5, result);

            result = RoundingHelper.Round(-22.25, 1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(-22, result);

            result = RoundingHelper.Round(-22.25, -1, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(-23, result);

            result = RoundingHelper.Round(-22.25, 5, RoundingHelper.Direction.AlwaysUp);
            Assert.AreEqual(-20, result);
        }

        [TestMethod]
        public void NearestRoundingTest()
        {
            // Direction.Nearest = midpoint away-from-zero on quotient (then * |multiple|)
            var result = RoundingHelper.Round(22.24, 0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(22.2, result);

            result = RoundingHelper.Round(22.25, 0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(22.26, 0.1, RoundingHelper.Direction.Nearest);
            Assert.AreEqual(22.3, result);

            result = RoundingHelper.Round(-22.25, -0.1, RoundingHelper.Direction.Nearest); // |multiple| = 0.1
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

            result = RoundingHelper.Round(-1555, -1000, RoundingHelper.Direction.Nearest); // |multiple| = 1000
            Assert.AreEqual(-2000, result);
        }

        [TestMethod]
        public void BigNumbersRoundingTest()
        {
            using(var p=new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                ws.Cells["A1"].Formula = "3627683204550360+555487768042629000";
                ws.Cells["A2"].Formula = "551879477094850000+71728430321556700000";
                ws.Calculate();
                Assert.AreEqual(559115451247179330D, (double)ws.Cells["A1"].Value);
                Assert.AreEqual(72280309798651552000D, (double)ws.Cells["A2"].Value);
            }
        } 
    }
}
