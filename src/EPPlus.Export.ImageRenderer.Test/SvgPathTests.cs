using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Xml;
using TypeConv = OfficeOpenXml.Utils.TypeConversion;

namespace TestProject1
{
    [TestClass]
    public sealed class SvgPathTests : TestBase
    {
        [TestMethod]
        public void Rect()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenPackage("svg/rect.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Rect);
                d.Text = "Rectangle Rectangle Rectangle Rectangle";
                d.TextAlignment = OfficeOpenXml.Drawing.eTextAlignment.Left;
                d.TextAnchoring = OfficeOpenXml.Drawing.eTextAnchoringType.Bottom;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("rect.svg", svg);
                SaveAndCleanup(p);
            }
        }
        [TestMethod]
        public void RoundRect()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", eShapeStyle.RoundRect);

                d.TextAlignment = OfficeOpenXml.Drawing.eTextAlignment.Left;
                d.TextAnchoring = OfficeOpenXml.Drawing.eTextAnchoringType.Bottom;
                d.Text = "Rectangle Rectangle Rectangle Rectangle";

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("Roundrect.svg", svg);
                SaveWorkbook("svgRoundRectdrawing.xlsx", p);
            }
        }
        [TestMethod]
        public void Triangle()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Triangle);
                d.Text = "Test";
                d.TextAlignment = OfficeOpenXml.Drawing.eTextAlignment.Center;
                d.TextAnchoring = OfficeOpenXml.Drawing.eTextAnchoringType.Center;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("Triangle.svg", svg);
                SaveWorkbook("svgTriangleDrawing.xlsx", p);
            }
        }
        [TestMethod]
        public void RightArrow()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.RightArrow);
                d.SetSize(100, 100);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("RightArrow.svg", svg);
                SaveWorkbook("svgRightArrowDrawing.xlsx", p);
            }
        }
        [TestMethod]
        public void SmileyFace()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.SmileyFace);
                d.SetSize(800, 800);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("SmileyFace.svg", svg);
                SaveWorkbook("svgSmileyFace.xlsx", p);
            }
        }
        [TestMethod]
        public void VerticalScroll()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.VerticalScroll);
                d.SetSize(800, 800);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("VerticalScroll.svg", svg);
                SaveWorkbook("svgVerticalScroll.xlsx", p);
            }
        }
        [TestMethod]
        public void CloudCallout()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.CloudCallout);
                d.SetSize(800, 800);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("CloudCallout.svg", svg);
                SaveWorkbook("svgCloudCallout.xlsx", p);
            }
        }
        [TestMethod]
        public void IrregularSeal2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.IrregularSeal2);
                d.SetSize(800, 800);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("IrregularSeal2.svg", svg);
                SaveWorkbook("IrregularSeal2.xlsx", p);
            }
        }
        [TestMethod]
        public void LightningBolt()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.LightningBolt);
                d.SetSize(800, 800);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("LightningBolt.svg", svg);
                SaveWorkbook("svgLightningBolt.xlsx", p);
            }
        }
        [TestMethod]
        public void FlowChartMagneticTape()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.FlowChartMagneticTape);
                d.SetSize(800, 800);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("FlowChartMagneticTape.svg", svg);
                SaveWorkbook("svgFlowChartMagneticTape.xlsx", p);
            }
        }
        [TestMethod]
        public void MathNotEqual()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.MathNotEqual);
                d.SetSize(800, 800);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("MathNotEqual.svg", svg);
                SaveWorkbook("svgMathNotEqual.xlsx", p);
            }
        }
        [TestMethod]
        public void Sun()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Sun);

                d.SetSize(800, 800);
                d.Fill.Color = Color.Orange;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("Sun.svg", svg);
                SaveWorkbook("svgSun.xlsx", p);
            }
        }
        [TestMethod]
        public void Ellipse()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Ellipse);

                d.SetSize(800, 800);
                d.Fill.Color = Color.Orange;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("Ellipse.svg", svg);
                SaveWorkbook("svgEllipse.xlsx", p);
            }
        }
        [TestMethod]
        public void Heart()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Heart);

                d.SetSize(800, 800);
                d.Fill.Color = Color.Red;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("Heart.svg", svg);
                SaveWorkbook("svgHeart.xlsx", p);
            }
        }
        [TestMethod]
        public void BevelRed()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Bevel);

                d.SetSize(804, 804);
                d.Fill.Color = Color.Red;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("bevelred.svg", svg);
                SaveWorkbook("svgBevelred.xlsx", p);
            }
        }
        [TestMethod]
        public void Bevel()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Bevel);

                d.SetSize(804, 804);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("bevel.svg", svg);
                SaveWorkbook("svgBevel.xlsx", p);
            }
        }

        [TestMethod]
        public void LeftBracket()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", eShapeStyle.LeftBracket);

                d.SetSize(804, 804);
                d.Fill.Style = eFillStyle.NoFill;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("LeftBracket.svg", svg);
                SaveWorkbook("LeftBracket.xlsx", p);
            }
        }
        [TestMethod]
        public void CalloutQuadArrow()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", eShapeStyle.QuadArrowCallout);

                d.SetSize(804, 200);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook("QuadArrowCallout.svg", svg);
                SaveWorkbook("QuadArrowCallout.xlsx", p);
            }
        }
        [TestMethod]
        public void ActionButtonHome()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonHome);

                d.SetSize(804, 804);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook($"svg\\ActionButtonHome.svg", svg);
                SaveWorkbook("ActionButtonHome.xlsx", p);
            }
        }
        [TestMethod]
        public void ActionButtonMovie()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", eShapeStyle.ActionButtonMovie);

                d.SetSize(804, 804);
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook($"svg\\ActionButtonMovie.svg", svg);
                SaveWorkbook("ActionButtonMovie.xlsx", p);
            }
        }
        [TestMethod]
        public void CustomPath()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage(@"svg\CustPath.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //var d = ws.Drawings[0].As.Shape;
                //Assert.AreEqual(1, d.CustomGeom.DrawingPaths.Count);
                //d.Textbox = "GetRectangle GetRectangle GetRectangle GetRectangle";
                //d.TextAlignment = OfficeOpenXml.Drawing.eTextAlignment.Left;
                //d.TextAnchoring = OfficeOpenXml.Drawing.eTextAnchoringType.Bottom;
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                //var svg = renderer.RenderDrawingToSvg(d);
                //SaveTextFileToWorkbook("CustomDrawing1.svg", svg);

                var d = ws.Drawings[1].As.Shape;
                var svg = renderer.RenderDrawingToSvg(d);
                SaveTextFileToWorkbook($"svg\\CustomDrawing2.svg", svg);

                //SaveWorkbook("svgdrawing.xlsx");
            }
        }


        [TestMethod]
        public void GenerateAllShapes()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Shapes");
                int y = 100, i = 1;
                foreach (eShapeStyle style in Enum.GetValues(typeof(eShapeStyle)))
                {
                    if (style == eShapeStyle.CustomShape) continue;
                    var shape = ws.Drawings.AddShape(style.ToString(), style);
                    shape.Text = style.ToString();
                    Assert.AreEqual(eDrawingType.Shape, shape.DrawingType);
                    shape.SetPosition(y, 100);
                    shape.SetSize(600, 600);
                    y += 700;
                    i++;
                }
                SaveWorkbook("shapes.xlsx", p);
            }
        }

        [TestMethod]
        public void TestShapes()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("margins.xlsx"))
            {
                var drawings = p.Workbook.Worksheets[0].Drawings;
                var textbody = drawings[0].As.Shape.TextBody;

                var insertCM = textbody.LeftInsert.Value * 0.0352777778;

                Assert.AreEqual(0.3d, insertCM, 0.00001);
            }
        }
        [TestMethod]
        public void GenerateSvgForGradientFilledShapes()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("GradientFillShapes.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                int ix = 1;

                //var d = ws.Drawings[7];
                //var svg = renderer.RenderDrawingToSvg(d);
                //File.WriteAllText($"c:\\temp\\{d.Name}-{ix}.svg", svg);

                foreach (var d in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(d);
                    SaveTextFileToWorkbook($"svg\\{d.Name}-{ix}.svg", svg);
                    ix++;
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForGradientRadialFilledShapes()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("GradiantRadial.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var d = ws.Drawings[5];
                //var svg = renderer.RenderDrawingToSvg(d);
                //File.WriteAllText($"c:\\temp\\{d.Name}-{5}.svg", svg);    

                int ix = 1;

                foreach (var d in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(d);
                    SaveTextFileToWorkbook($"svg\\{d.Name}-{ix}.svg", svg);
                    ix++;
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForPatternFilledShapes()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("PatternFills.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                int ix = 1;

                foreach (ExcelShape d in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(d);
                    SaveTextFileToWorkbook($"svg\\Pattern-{d.Fill.PatternFill.PatternType}-{ix}.svg", svg);
                    ix++;
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForBlipFillShapes()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("BlipFills.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                int ix = 1;

                foreach (ExcelShape d in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(d);
                    SaveSvg($"Blip{ix}.svg", svg);
                    ix++;
                }
            }
        }


        [TestMethod]
        public void GenerateSvgForCircle()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("GradientRadialVerifyCircle.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var svg = renderer.RenderDrawingToSvg(ws.Drawings[1]);
                //File.WriteAllText($"c:\\temp\\ChartForSvg{1}.svg", svg);

                int ix = 1;

                foreach (ExcelShape d in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(d);
                    SaveTextFileToWorkbook($"svg\\DrawForSvg{ix}.svg", svg);
                    ix++;
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                //var ix = 3;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\ChartForSvg_ind{ix++}.svg", svg);

                var ix = 0;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\ChartForSvg{ix++}.svg", svg);
                }
            }
        }

        [TestMethod]
        public void GenerateSimplestChart()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("SimplestChart.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\SimplestChartTitle.svg", svg);
            }
        }


        [TestMethod]
        public void OpenRightAligned()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("SimpleChartRightAlign.xlsx"))
            {
                var c = p.Workbook.Worksheets[0].Drawings[0];

                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\SimplestChartRightAlign.svg", svg);
            }
        }

        [TestMethod]
        public void GenerateShapeCenteredParagraph()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenPackage("ShapeTestCentered.xlsx",true))
            {
                var sheet = p.Workbook.Worksheets.Add("ShapeSheet");

                var _currentShape = sheet.Drawings.AddShape("CubeTest", eShapeStyle.Cube);

                _currentShape.SetPixelWidth(300d);
                _currentShape.SetPixelHeight(300d);

                _currentShape.Fill.Style = eFillStyle.SolidFill;
                _currentShape.Fill.Color = System.Drawing.Color.BlueViolet;
                _currentShape.Font.Color = System.Drawing.Color.Goldenrod;

                _currentShape.TextBody.TopInsert = 0;
                _currentShape.TextBody.BottomInsert = 0;
                _currentShape.TextBody.RightInsert = 0;
                _currentShape.TextBody.LeftInsert = 0;

                var para1 = _currentShape.TextBody.Paragraphs.Add("TextBox\r\na");
                //var test = _currentShape.TextBody.AnchorCenter;

                para1.LeftMargin = 5;
                _currentShape.TextBody.TopInsert = 10;

                var para2 = _currentShape.TextBody.Paragraphs.Add("TextBox2");
                para2.TextRuns[0].FontItalic = true;
                para2.TextRuns[0].FontBold = true;
                para2.TextRuns.Add("ra underline").FontUnderLine = eUnderLineType.Dash;
                para2.TextRuns.Add("La Strike").FontStrike = eStrikeType.Single;
                var tRun1 = para2.TextRuns.Add("Goudy size 16");
                tRun1.SetFromFont("Goudy Stout", 16);

                _currentShape.TextBody.Paragraphs[0].HorizontalAlignment = eTextAlignment.Center;

                _currentShape.TextAnchoring = eTextAnchoringType.Top;

                //var smiley = "\ud83d\ude03";

                tRun1.Fill.Color = System.Drawing.Color.IndianRed;
                var tRun2 = para2.TextRuns.Add("SvgSize 24");
                tRun2.FontSize = 24;

                //_currentShape.TextAnchoring = eTextAnchoringType.Center;

                _currentShape.TextBody.HorizontalTextOverflow = eTextHorizontalOverflow.Clip;
                _currentShape.TextBody.VerticalTextOverflow = eTextVerticalOverflow.Clip;

                
                //SetFillColor(_currentShape.Fill, txtFillColor.Text);
                //SetFillColor(_currentShape.Border.Fill, txtBorderColor.Text);

                var aFont = _currentShape.Font;
                var paragraph0 = _currentShape.TextBody.Paragraphs[0];

                _currentShape.GetSizeInPixels(out int testWidth, out int testHeight);

                SaveAndCleanup(p);
            }
        }

        [TestMethod]
        public void GenerateSvgForLineCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("LineChartRenderTest.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var ix = 2;
                //var c = ws.Drawings[ix];
                //var svg = renderer.RenderDrawingToSvg(c);
                //SaveTextFileToWorkbook($"svg\\LineChartForSvg_Single{ix++}.svg", svg);
                var ix = 1;
                foreach (ExcelChart c in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(c);
                    SaveTextFileToWorkbook($"svg\\LineChartForSvg{ix++}.svg", svg);
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForCharts_sheet2()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[1];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var ix = 1;
                var c = ws.Drawings[ix];
                var svg = renderer.RenderDrawingToSvg(c);
                SaveTextFileToWorkbook($"svg\\ChartForSvg_sheet2_{ix++}.svg", svg);
                //var ix = 1;
                //foreach (ExcelChart c in ws.Drawings)
                //{
                //    var svg = renderer.RenderDrawingToSvg(c);
                //    SaveTextFileToWorkbook($"svg\\ChartForSvg{ix++}.svg", svg);
                //}
            }
        }

        [TestMethod]
        public void CreateChartsWithDifferentSize()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = OpenPackage("ChartWithDifferentSizes.xlsx", true))
            {
                var ws = p.Workbook.Worksheets.Add("Chart1");

                LoadItemData(ws);
                var chart1 = ws.Drawings.AddChart("chart1", eChartType.Line);
                chart1.Title.Text = "This is a very long title that should be wrapped into multiple lines.";
                chart1.Title.Font.Size = 32;
                chart1.Series.Add(ws.Cells["O2:O11"], ws.Cells["N2:N10"]);
                chart1.SetPosition(2, 0, 1, 0);
                chart1.SetSize(400, 400);

                var chart2 = ws.Drawings.AddChart("chart2", eChartType.ColumnClustered);
                chart2.Title.Text = "This is a very long title that should be wrapped into multiple lines.";
                chart2.Title.Font.Size = 32;
                chart2.Series.Add(ws.Cells["O2:O11"], ws.Cells["N2:N10"]);
                chart2.SetPosition(25, 0, 1, 0);
                chart2.SetSize(400, 800);

                var chart3 = ws.Drawings.AddChart("chart3", eChartType.BarClustered);
                chart3.Title.Text = "This is a very long title that should be wrapped into multiple lines.";
                chart3.Title.Font.Size = 32;
                chart3.Series.Add(ws.Cells["O2:O11"], ws.Cells["N2:N10"]);
                chart3.SetPosition(25, 0, 8, 0);
                chart3.SetSize(800, 400);

                SaveAndCleanup(p);
            }
        }

    }
}
