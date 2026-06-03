using EPPlus.DrawingRenderer;
using EPPlus.DrawingRenderer.ShapeDefinitions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;

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

                PresetShapeDefinitions.ShapeDefinitions[(ShapeStyle)shape.Style].Calculate(shape._width, shape._height, shape.TextBody.TextAutofit == eTextAutofit.ShapeAutofit, null, null);
                await p.SaveAsAsync("c:\\temp\\rect.xlsx");
            }
        }

        [TestMethod]
        public void MathMultiplyTest()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPLUS");
            var shDef = PresetShapeDefinitions.ShapeDefinitions[ShapeStyle.MathMultiply];

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws");
                var shape = ws.Drawings.AddShape("mMult", eShapeStyle.MathMultiply);
                shDef.Calculate(shape._width, shape._height, shape.TextBody.TextAutofit == eTextAutofit.ShapeAutofit, null, null);
            }
        }

        [TestMethod]
        public void CurvedDownArrow()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPLUS");
            var shDef = PresetShapeDefinitions.ShapeDefinitions[ShapeStyle.CurvedDownArrow];

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws");
                var shape = ws.Drawings.AddShape("mMult", eShapeStyle.CurvedDownArrow);
                shape.SetSize(100, 100);

                shDef.Calculate(shape._width, shape._height, shape.TextBody.TextAutofit == eTextAutofit.ShapeAutofit, null, null);


                //shDef._calculatedValues

            }
        }
        [TestMethod]
        public void BlockArc()
        {
            ExcelPackage.License.SetNonCommercialPersonal("EPPLUS");
            var shDef = PresetShapeDefinitions.ShapeDefinitions[ShapeStyle.BlockArc];

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ws");
                var shape = ws.Drawings.AddShape("mMult", eShapeStyle.BlockArc);
                shape.SetSize(100, 100);

                shDef.Calculate(shape._width, shape._height, shape.TextBody.TextAutofit == eTextAutofit.ShapeAutofit, null, null);
                //shDef._calculatedValues
            }
        }
    }
}
