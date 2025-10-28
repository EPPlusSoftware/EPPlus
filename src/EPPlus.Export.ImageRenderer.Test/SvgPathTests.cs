using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using System.Diagnostics;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Xml;
using TypeConv = OfficeOpenXml.Utils.TypeConversion;
using EPPlus.Fonts.OpenType;

namespace TestProject1
{
    [TestClass]
    public sealed class SvgPathTests : TestBase
    {
        [TestMethod]
        public void Rect()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                var d = ws.Drawings.AddShape("Shape1", OfficeOpenXml.Drawing.eShapeStyle.Rect);
                d.Text = "Rectangle Rectangle Rectangle Rectangle";
                d.TextAlignment = OfficeOpenXml.Drawing.eTextAlignment.Left;
                d.TextAnchoring = OfficeOpenXml.Drawing.eTextAnchoringType.Bottom;
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                var svg = renderer.RenderDrawingToSvg(d);
                File.WriteAllText("c:\\temp\\rect.svg", svg);
                p.SaveAs("c:\\temp\\svgdrawing.xlsx");
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
                File.WriteAllText("c:\\temp\\Roundrect.svg", svg);
                p.SaveAs("c:\\temp\\svgRoundRectdrawing.xlsx");
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
                File.WriteAllText("c:\\temp\\Triangle.svg", svg);
                p.SaveAs("c:\\temp\\svgTriangleDrawing.xlsx");
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
                File.WriteAllText("c:\\temp\\RightArrow.svg", svg);
                p.SaveAs("c:\\temp\\svgRightArrowDrawing.xlsx");
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
                File.WriteAllText("c:\\temp\\SmileyFace.svg", svg);
                p.SaveAs("c:\\temp\\svgSmileyFace.xlsx");
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
                File.WriteAllText("c:\\temp\\VerticalScroll.svg", svg);
                p.SaveAs("c:\\temp\\svgVerticalScroll.xlsx");
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
                File.WriteAllText("c:\\temp\\CloudCallout.svg", svg);
                p.SaveAs("c:\\temp\\svgCloudCallout.xlsx");
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
                File.WriteAllText("c:\\temp\\IrregularSeal2.svg", svg);
                p.SaveAs("c:\\temp\\IrregularSeal2.xlsx");
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
                File.WriteAllText("c:\\temp\\LightningBolt.svg", svg);
                p.SaveAs("c:\\temp\\svgLightningBolt.xlsx");
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
                File.WriteAllText("c:\\temp\\FlowChartMagneticTape.svg", svg);
                p.SaveAs("c:\\temp\\svgFlowChartMagneticTape.xlsx");
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
                File.WriteAllText("c:\\temp\\MathNotEqual.svg", svg);
                p.SaveAs("c:\\temp\\svgMathNotEqual.xlsx");
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
                File.WriteAllText("c:\\temp\\Sun.svg", svg);
                p.SaveAs("c:\\temp\\svgSun.xlsx");
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
                File.WriteAllText("c:\\temp\\Ellipse.svg", svg);
                p.SaveAs("c:\\temp\\svgEllipse.xlsx");
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
                File.WriteAllText("c:\\temp\\Heart.svg", svg);
                p.SaveAs("c:\\temp\\svgHeart.xlsx");
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
                File.WriteAllText("c:\\temp\\bevelred.svg", svg);
                p.SaveAs("c:\\temp\\svgBevelred.xlsx");
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
                File.WriteAllText("c:\\temp\\bevel.svg", svg);
                p.SaveAs("c:\\temp\\svgBevel.xlsx");
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
                File.WriteAllText("c:\\temp\\LeftBracket.svg", svg);
                p.SaveAs("c:\\temp\\LeftBracket.xlsx");
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
                File.WriteAllText("c:\\temp\\QuadArrowCallout.svg", svg);
                p.SaveAs("c:\\temp\\QuadArrowCallout.xlsx");
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
                File.WriteAllText("c:\\temp\\ActionButtonHome.svg", svg);
                p.SaveAs("c:\\temp\\ActionButtonHome.xlsx");
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
                File.WriteAllText("c:\\temp\\ActionButtonMovie.svg", svg);
                p.SaveAs("c:\\temp\\ActionButtonMovie.xlsx");
            }
        }
        [TestMethod]
        public void CustomPath()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage(@"c:\temp\CustPath.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                //var d = ws.Drawings[0].As.Shape;
                //Assert.AreEqual(1, d.CustomGeom.DrawingPaths.Count);
                //d.Text = "Rectangle Rectangle Rectangle Rectangle";
                //d.TextAlignment = OfficeOpenXml.Drawing.eTextAlignment.Left;
                //d.TextAnchoring = OfficeOpenXml.Drawing.eTextAnchoringType.Bottom;
                var renderer = new EPPlusImageRenderer.ImageRenderer();

                //var svg = renderer.RenderDrawingToSvg(d);
                //File.WriteAllText("c:\\temp\\CustomDrawing1.svg", svg);

                var d = ws.Drawings[1].As.Shape;
                var svg = renderer.RenderDrawingToSvg(d);
                File.WriteAllText("c:\\temp\\CustomDrawing2.svg", svg);

                //p.SaveAs("c:\\temp\\svgdrawing.xlsx");
            }
        }
        [TestMethod]
        public void ReadSvgFilesAndRename()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage(@"c:\Users\Janne\Pictures\MacroToGetAllAutoShapes.xlsm"))
            {
                var ws = p.Workbook.Worksheets[0];
                var adjPoints = LoadPoints();
                var sb = new StringBuilder();
                for (int i = 0; i < ws.Drawings.Count; i++)
                {
                    var d = ws.Drawings[i].As.Shape;
                    var file1 = $"c:\\Users\\Janne\\Pictures\\{d.Text}.svg";
                    var file2 = $"c:\\Users\\Janne\\Pictures\\CorrectFileNames\\{d.As.Shape.Style}.svg";
                    if (File.Exists(file1) && !File.Exists(file2))
                    {
                        File.Copy(file1, file2, true);
                    }
                    if (adjPoints.TryGetValue(d.Text, out List<double> list))
                    {
                        //sb.Append($"{{eShapeStyle.{d.Style}, new List<double>(){{");
                        sb.Append($"{{eShapeStyle.{d.Style},new Dictionary<string, ShapeGuidePoint> {{ ");
                        int x = 1;
                        foreach (var ap in list)
                        {
                            if (list.Count == 1)
                            {
                                sb.Append($"{{\"adj\", {ap * 100000}}},");
                            }   
                            else
                            {
                                sb.Append($"{{\"adj{x++}\", {ap * 100000}}},");
                            }
                            //sb.Append($"{ap.ToString(CultureInfo.InvariantCulture)},");
                        }
                        sb.Length--;
                        sb.Append("}},\r\n");
                    }
                }
                File.WriteAllText("c:\\temp\\adjstCode.txt", sb.ToString());
            }
        }

        private Dictionary<string,List<double>> LoadPoints()
        {
            var lines = File.ReadAllLines("c:\\temp\\adjust.txt");            
            var d = new Dictionary<string, List<double>>();
            List<double>? list;
            var hs = new HashSet<string>();
            var pc = "";

            foreach (var l in lines)
            {
                var cols = l.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (pc == cols[0] || hs.Contains(cols[0]) == false)
                {
                    if (!d.TryGetValue(cols[0], out list))
                    {
                        list = new List<double>();
                        d.Add(cols[0], list);
                    }
                    list.Add(double.Parse(cols[1]));
                }
                pc = cols[0];
            }
            return d;
        }

        [TestMethod]
        public void GetCorrectValuesForEnum()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            using (var p = new ExcelPackage())
            {
                var theme = p.Workbook.ThemeManager.GetOrCreateTheme();
                var fc = TypeConv.ColorConverter.GetThemeColor(theme.ColorScheme.Accent1);

                for (int i = 10; i < 100; i++)
                {
                    for (int o1 = 0; o1 <= 100; o1 += 10)
                    {
                        for (int o2 = 0; o2 <= 100; o2 += 10)
                        {
                            for (int s = 20; s <= 80; s++)
                            {
                                var c = TypeConv.ColorConverter.ApplySatMod(fc, s / 100D, o1 / 100);
                                c = TypeConv.ColorConverter.ApplyLumMod(c, (double)i / 100, (double)o2 / 100);

                                if (c.R == 0x43 && c.G == 0x7F && c.B == 0x9a)
                                {
                                    Debug.WriteLine(i, $"Color:{c.Name},Lum={i}, LOff={o2}, Sat={s}, SatOff={o1}");
                                }
                                //if (c.R == 0x73 && c.G == 0xA0)
                                //{
                                //    Debug.WriteLine(i, $"Color:{c.Name},Lum={i}, LOff={o2}, Sat={s}, SatOff={o1}");
                                //}
                            }
                        }
                    }
                }

                var darken = TypeConv.ColorConverter.ApplyLumMod(fc, 0.10, 0.15);   //#0D3A4E
                var darkenless = TypeConv.ColorConverter.ApplyLumMod(fc, 0.3, 0.15); //#114D69
                var lightenless = TypeConv.ColorConverter.ApplyTint(fc, 0.25); //#73A0B4
                var light = TypeConv.ColorConverter.ApplyTint(fc, 0.60); //#437F9B
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
                p.SaveAs("c:\\temp\\shapes.xlsx");
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
                    File.WriteAllText($"c:\\temp\\{d.Name}-{ix}.svg", svg);
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
                    File.WriteAllText($"c:\\temp\\{d.Name}-{ix}.svg", svg);
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
                    File.WriteAllText($"c:\\temp\\Pattern-{d.Fill.PatternFill.PatternType}-{ix}.svg", svg);
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
                    File.WriteAllText($"c:\\temp\\Blip{ix}.svg", svg);
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
                    File.WriteAllText($"c:\\temp\\DrawForSvg{ix}.svg", svg);
                    ix++;
                }
            }
        }
        [TestMethod]
        public void GenerateSvgForCharts()
        {
            ExcelPackage.License.SetNonCommercialOrganization("EPPlus Project");
            //TextData.FontDirectories.Add("c:\\fonts");
            using (var p = OpenTemplatePackage("ChartForSvg.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var renderer = new EPPlusImageRenderer.ImageRenderer();
                //var svg = renderer.RenderDrawingToSvg(ws.Drawings[1]);
                //File.WriteAllText($"c:\\temp\\ChartForSvg{1}.svg", svg);
                int ix = 1;
                foreach (ExcelChart d in ws.Drawings)
                {
                    var svg = renderer.RenderDrawingToSvg(d);
                    File.WriteAllText($"c:\\temp\\ChartForSvg{ix}.svg", svg);
                    ix++;
                }
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
