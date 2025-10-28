using EPPlusImageRenderer;
using EPPlusImageRenderer.ShapeDefinitions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using OfficeOpenXml.Utils;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;

namespace TestProject1
{
    [TestClass]
    public sealed class PresetShapeDefinitionTests
    {
        [TestMethod]
        public async Task LoadPreset()
        {
            var psd = PresetShapeDefinitions.ShapeDefinitions;
            Assert.AreEqual(187, psd.Count);

            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var shape = ws.Drawings.AddShape("Rect1", OfficeOpenXml.Drawing.eShapeStyle.Rect);

                PresetShapeDefinitions.ShapeDefinitions[shape.Style].Calculate(shape);
                p.SaveAs("c:\\temp\\rect.xlsx");
            }
        }

        [TestMethod]
        public void MathMultiplyTest()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPLUS");
            var shDef = PresetShapeDefinitions.ShapeDefinitions[eShapeStyle.MathMultiply];

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws");
                var drawing = ws.Drawings.AddShape("mMult", eShapeStyle.MathMultiply);
                shDef.Calculate(drawing);
            }
        }

        [TestMethod]
        public void CurvedDownArrow()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPLUS");
            var shDef = PresetShapeDefinitions.ShapeDefinitions[eShapeStyle.CurvedDownArrow];

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws");
                var drawing = ws.Drawings.AddShape("mMult", eShapeStyle.CurvedDownArrow);
                drawing.SetSize(100, 100);

                shDef.Calculate(drawing);


                //shDef._calculatedValues

            }
        }
        [TestMethod]
        public void BlockArc()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPLUS");
            var shDef = PresetShapeDefinitions.ShapeDefinitions[eShapeStyle.BlockArc];

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws");
                var drawing = ws.Drawings.AddShape("mMult", eShapeStyle.BlockArc);
                drawing.SetSize(100, 100);

                shDef.Calculate(drawing);
                //shDef._calculatedValues
            }
        }
    }
}
