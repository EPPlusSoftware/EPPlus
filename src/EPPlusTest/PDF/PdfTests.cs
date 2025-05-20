using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.PDF;
using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.PdfPageSettings.PdfPageSizes;

namespace EPPlusTest.PDF
{
    [TestClass]
    public class PdfTests : TestBase
    {
        [TestMethod]
        public void TestWritePdf()
        {
            using var p = OpenTemplatePackage("PDFTest.xlsx");
            var ws = p.Workbook.Worksheets[0];

            ExcelPdf pedeef = new ExcelPdf();

            pedeef.CreatePdf("c:\\epplustest\\pdf\\FullPageTest2.pdf", ws);

        }






        [TestMethod]
        public void ChartToSVG()
        {
            string chart = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">\r\n  <c:chart>\r\n    <c:plotArea>\r\n      <c:barChart>\r\n        <c:ser>\r\n          <c:cat>\r\n            <c:strRef>\r\n              <c:strCache>\r\n                <c:pt idx=\"0\"><c:v>Apples</c:v></c:pt>\r\n                <c:pt idx=\"1\"><c:v>Oranges</c:v></c:pt>\r\n                <c:pt idx=\"2\"><c:v>Bananas</c:v></c:pt>\r\n              </c:strCache>\r\n            </c:strRef>\r\n          </c:cat>\r\n          <c:val>\r\n            <c:numRef>\r\n              <c:numCache>\r\n                <c:pt idx=\"0\"><c:v>10</c:v></c:pt>\r\n                <c:pt idx=\"1\"><c:v>20</c:v></c:pt>\r\n                <c:pt idx=\"2\"><c:v>15</c:v></c:pt>\r\n              </c:numCache>\r\n            </c:numRef>\r\n          </c:val>\r\n        </c:ser>\r\n      </c:barChart>\r\n    </c:plotArea>\r\n  </c:chart>\r\n</c:chartSpace>";
            string chart2 = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"\r\n    xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"\r\n    xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\r\n    xmlns:c16r2=\"http://schemas.microsoft.com/office/drawing/2015/06/chart\">\r\n    <c:date1904 val=\"0\" />\r\n    <c:lang val=\"en-US\" />\r\n    <c:roundedCorners val=\"0\" />\r\n    <mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">\r\n        <mc:Choice Requires=\"c14\"\r\n            xmlns:c14=\"http://schemas.microsoft.com/office/drawing/2007/8/2/chart\">\r\n            <c14:style val=\"102\" />\r\n        </mc:Choice>\r\n        <mc:Fallback>\r\n            <c:style val=\"2\" />\r\n        </mc:Fallback>\r\n    </mc:AlternateContent>\r\n    <c:chart>\r\n        <c:title>\r\n            <c:overlay val=\"0\" />\r\n            <c:spPr>\r\n                <a:noFill />\r\n                <a:ln>\r\n                    <a:noFill />\r\n                </a:ln>\r\n                <a:effectLst />\r\n            </c:spPr>\r\n            <c:txPr>\r\n                <a:bodyPr rot=\"0\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\" vert=\"horz\"\r\n                    wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\" />\r\n                <a:lstStyle />\r\n                <a:p>\r\n                    <a:pPr>\r\n                        <a:defRPr sz=\"1600\" b=\"1\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\"\r\n                            baseline=\"0\">\r\n                            <a:solidFill>\r\n                                <a:schemeClr val=\"tx1\">\r\n                                    <a:lumMod val=\"65000\" />\r\n                                    <a:lumOff val=\"35000\" />\r\n                                </a:schemeClr>\r\n                            </a:solidFill>\r\n                            <a:latin typeface=\"+mn-lt\" />\r\n                            <a:ea typeface=\"+mn-ea\" />\r\n                            <a:cs typeface=\"+mn-cs\" />\r\n                        </a:defRPr>\r\n                    </a:pPr>\r\n                    <a:endParaRPr lang=\"en-SE\" />\r\n                </a:p>\r\n            </c:txPr>\r\n        </c:title>\r\n        <c:autoTitleDeleted val=\"0\" />\r\n        <c:view3D>\r\n            <c:rotX val=\"15\" />\r\n            <c:rotY val=\"20\" />\r\n            <c:depthPercent val=\"100\" />\r\n            <c:rAngAx val=\"1\" />\r\n        </c:view3D>\r\n        <c:floor>\r\n            <c:thickness val=\"0\" />\r\n            <c:spPr>\r\n                <a:noFill />\r\n                <a:ln>\r\n                    <a:noFill />\r\n                </a:ln>\r\n                <a:effectLst />\r\n                <a:sp3d />\r\n            </c:spPr>\r\n        </c:floor>\r\n        <c:sideWall>\r\n            <c:thickness val=\"0\" />\r\n            <c:spPr>\r\n                <a:noFill />\r\n                <a:ln>\r\n                    <a:noFill />\r\n                </a:ln>\r\n                <a:effectLst />\r\n                <a:sp3d />\r\n            </c:spPr>\r\n        </c:sideWall>\r\n        <c:backWall>\r\n            <c:thickness val=\"0\" />\r\n            <c:spPr>\r\n                <a:noFill />\r\n                <a:ln>\r\n                    <a:noFill />\r\n                </a:ln>\r\n                <a:effectLst />\r\n                <a:sp3d />\r\n            </c:spPr>\r\n        </c:backWall>\r\n        <c:plotArea>\r\n            <c:layout />\r\n            <c:bar3DChart>\r\n                <c:barDir val=\"col\" />\r\n                <c:grouping val=\"clustered\" />\r\n                <c:varyColors val=\"0\" />\r\n                <c:ser>\r\n                    <c:idx val=\"0\" />\r\n                    <c:order val=\"0\" />\r\n                    <c:tx>\r\n                        <c:strRef>\r\n                            <c:f>Sheet1!$B$1</c:f>\r\n                            <c:strCache>\r\n                                <c:ptCount val=\"1\" />\r\n                                <c:pt idx=\"0\">\r\n                                    <c:v>Amount</c:v>\r\n                                </c:pt>\r\n                            </c:strCache>\r\n                        </c:strRef>\r\n                    </c:tx>\r\n                    <c:spPr>\r\n                        <a:gradFill rotWithShape=\"1\">\r\n                            <a:gsLst>\r\n                                <a:gs pos=\"0\">\r\n                                    <a:schemeClr val=\"accent1\">\r\n                                        <a:satMod val=\"103000\" />\r\n                                        <a:lumMod val=\"102000\" />\r\n                                        <a:tint val=\"94000\" />\r\n                                    </a:schemeClr>\r\n                                </a:gs>\r\n                                <a:gs pos=\"50000\">\r\n                                    <a:schemeClr val=\"accent1\">\r\n                                        <a:satMod val=\"110000\" />\r\n                                        <a:lumMod val=\"100000\" />\r\n                                        <a:shade val=\"100000\" />\r\n                                    </a:schemeClr>\r\n                                </a:gs>\r\n                                <a:gs pos=\"100000\">\r\n                                    <a:schemeClr val=\"accent1\">\r\n                                        <a:lumMod val=\"99000\" />\r\n                                        <a:satMod val=\"120000\" />\r\n                                        <a:shade val=\"78000\" />\r\n                                    </a:schemeClr>\r\n                                </a:gs>\r\n                            </a:gsLst>\r\n                            <a:lin ang=\"5400000\" scaled=\"0\" />\r\n                        </a:gradFill>\r\n                        <a:ln>\r\n                            <a:noFill />\r\n                        </a:ln>\r\n                        <a:effectLst>\r\n                            <a:outerShdw blurRad=\"57150\" dist=\"19050\" dir=\"5400000\" algn=\"ctr\"\r\n                                rotWithShape=\"0\">\r\n                                <a:srgbClr val=\"000000\">\r\n                                    <a:alpha val=\"63000\" />\r\n                                </a:srgbClr>\r\n                            </a:outerShdw>\r\n                        </a:effectLst>\r\n                        <a:sp3d />\r\n                    </c:spPr>\r\n                    <c:invertIfNegative val=\"0\" />\r\n                    <c:cat>\r\n                        <c:strRef>\r\n                            <c:f>Sheet1!$A$2:$A$4</c:f>\r\n                            <c:strCache>\r\n                                <c:ptCount val=\"3\" />\r\n                                <c:pt idx=\"0\">\r\n                                    <c:v>PDF</c:v>\r\n                                </c:pt>\r\n                                <c:pt idx=\"1\">\r\n                                    <c:v>??????</c:v>\r\n                                </c:pt>\r\n                                <c:pt idx=\"2\">\r\n                                    <c:v>Profit</c:v>\r\n                                </c:pt>\r\n                            </c:strCache>\r\n                        </c:strRef>\r\n                    </c:cat>\r\n                    <c:val>\r\n                        <c:numRef>\r\n                            <c:f>Sheet1!$B$2:$B$4</c:f>\r\n                            <c:numCache>\r\n                                <c:formatCode>General</c:formatCode>\r\n                                <c:ptCount val=\"3\" />\r\n                                <c:pt idx=\"0\">\r\n                                    <c:v>1</c:v>\r\n                                </c:pt>\r\n                                <c:pt idx=\"1\">\r\n                                    <c:v>500</c:v>\r\n                                </c:pt>\r\n                                <c:pt idx=\"2\">\r\n                                    <c:v>1000</c:v>\r\n                                </c:pt>\r\n                            </c:numCache>\r\n                        </c:numRef>\r\n                    </c:val>\r\n                    <c:extLst>\r\n                        <c:ext uri=\"{C3380CC4-5D6E-409C-BE32-E72D297353CC}\"\r\n                            xmlns:c16=\"http://schemas.microsoft.com/office/drawing/2014/chart\">\r\n                            <c16:uniqueId val=\"{00000000-166E-40A6-91BB-6F261A06D3D1}\" />\r\n                        </c:ext>\r\n                    </c:extLst>\r\n                </c:ser>\r\n                <c:dLbls>\r\n                    <c:showLegendKey val=\"0\" />\r\n                    <c:showVal val=\"0\" />\r\n                    <c:showCatName val=\"0\" />\r\n                    <c:showSerName val=\"0\" />\r\n                    <c:showPercent val=\"0\" />\r\n                    <c:showBubbleSize val=\"0\" />\r\n                </c:dLbls>\r\n                <c:gapWidth val=\"150\" />\r\n                <c:shape val=\"box\" />\r\n                <c:axId val=\"922703887\" />\r\n                <c:axId val=\"922701007\" />\r\n                <c:axId val=\"0\" />\r\n            </c:bar3DChart>\r\n            <c:catAx>\r\n                <c:axId val=\"922703887\" />\r\n                <c:scaling>\r\n                    <c:orientation val=\"minMax\" />\r\n                </c:scaling>\r\n                <c:delete val=\"0\" />\r\n                <c:axPos val=\"b\" />\r\n                <c:numFmt formatCode=\"General\" sourceLinked=\"1\" />\r\n                <c:majorTickMark val=\"none\" />\r\n                <c:minorTickMark val=\"none\" />\r\n                <c:tickLblPos val=\"nextTo\" />\r\n                <c:spPr>\r\n                    <a:noFill />\r\n                    <a:ln w=\"12700\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\r\n                        <a:solidFill>\r\n                            <a:schemeClr val=\"tx1\">\r\n                                <a:lumMod val=\"15000\" />\r\n                                <a:lumOff val=\"85000\" />\r\n                            </a:schemeClr>\r\n                        </a:solidFill>\r\n                        <a:round />\r\n                    </a:ln>\r\n                    <a:effectLst />\r\n                </c:spPr>\r\n                <c:txPr>\r\n                    <a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\"\r\n                        vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\" />\r\n                    <a:lstStyle />\r\n                    <a:p>\r\n                        <a:pPr>\r\n                            <a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\"\r\n                                baseline=\"0\">\r\n                                <a:solidFill>\r\n                                    <a:schemeClr val=\"tx1\">\r\n                                        <a:lumMod val=\"65000\" />\r\n                                        <a:lumOff val=\"35000\" />\r\n                                    </a:schemeClr>\r\n                                </a:solidFill>\r\n                                <a:latin typeface=\"+mn-lt\" />\r\n                                <a:ea typeface=\"+mn-ea\" />\r\n                                <a:cs typeface=\"+mn-cs\" />\r\n                            </a:defRPr>\r\n                        </a:pPr>\r\n                        <a:endParaRPr lang=\"en-SE\" />\r\n                    </a:p>\r\n                </c:txPr>\r\n                <c:crossAx val=\"922701007\" />\r\n                <c:crosses val=\"autoZero\" />\r\n                <c:auto val=\"1\" />\r\n                <c:lblAlgn val=\"ctr\" />\r\n                <c:lblOffset val=\"100\" />\r\n                <c:noMultiLvlLbl val=\"0\" />\r\n            </c:catAx>\r\n            <c:valAx>\r\n                <c:axId val=\"922701007\" />\r\n                <c:scaling>\r\n                    <c:orientation val=\"minMax\" />\r\n                </c:scaling>\r\n                <c:delete val=\"0\" />\r\n                <c:axPos val=\"l\" />\r\n                <c:majorGridlines>\r\n                    <c:spPr>\r\n                        <a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\r\n                            <a:solidFill>\r\n                                <a:schemeClr val=\"tx1\">\r\n                                    <a:lumMod val=\"15000\" />\r\n                                    <a:lumOff val=\"85000\" />\r\n                                </a:schemeClr>\r\n                            </a:solidFill>\r\n                            <a:round />\r\n                        </a:ln>\r\n                        <a:effectLst />\r\n                    </c:spPr>\r\n                </c:majorGridlines>\r\n                <c:numFmt formatCode=\"General\" sourceLinked=\"1\" />\r\n                <c:majorTickMark val=\"none\" />\r\n                <c:minorTickMark val=\"none\" />\r\n                <c:tickLblPos val=\"nextTo\" />\r\n                <c:spPr>\r\n                    <a:noFill />\r\n                    <a:ln>\r\n                        <a:noFill />\r\n                    </a:ln>\r\n                    <a:effectLst />\r\n                </c:spPr>\r\n                <c:txPr>\r\n                    <a:bodyPr rot=\"-60000000\" spcFirstLastPara=\"1\" vertOverflow=\"ellipsis\"\r\n                        vert=\"horz\" wrap=\"square\" anchor=\"ctr\" anchorCtr=\"1\" />\r\n                    <a:lstStyle />\r\n                    <a:p>\r\n                        <a:pPr>\r\n                            <a:defRPr sz=\"900\" b=\"0\" i=\"0\" u=\"none\" strike=\"noStrike\" kern=\"1200\"\r\n                                baseline=\"0\">\r\n                                <a:solidFill>\r\n                                    <a:schemeClr val=\"tx1\">\r\n                                        <a:lumMod val=\"65000\" />\r\n                                        <a:lumOff val=\"35000\" />\r\n                                    </a:schemeClr>\r\n                                </a:solidFill>\r\n                                <a:latin typeface=\"+mn-lt\" />\r\n                                <a:ea typeface=\"+mn-ea\" />\r\n                                <a:cs typeface=\"+mn-cs\" />\r\n                            </a:defRPr>\r\n                        </a:pPr>\r\n                        <a:endParaRPr lang=\"en-SE\" />\r\n                    </a:p>\r\n                </c:txPr>\r\n                <c:crossAx val=\"922703887\" />\r\n                <c:crosses val=\"autoZero\" />\r\n                <c:crossBetween val=\"between\" />\r\n            </c:valAx>\r\n            <c:spPr>\r\n                <a:noFill />\r\n                <a:ln>\r\n                    <a:noFill />\r\n                </a:ln>\r\n                <a:effectLst />\r\n            </c:spPr>\r\n        </c:plotArea>\r\n        <c:plotVisOnly val=\"1\" />\r\n        <c:dispBlanksAs val=\"gap\" />\r\n        <c:extLst>\r\n            <c:ext uri=\"{56B9EC1D-385E-4148-901F-78D8002777C0}\"\r\n                xmlns:c16r3=\"http://schemas.microsoft.com/office/drawing/2017/03/chart\">\r\n                <c16r3:dataDisplayOptions16>\r\n                    <c16r3:dispNaAsBlank val=\"1\" />\r\n                </c16r3:dataDisplayOptions16>\r\n            </c:ext>\r\n        </c:extLst>\r\n        <c:showDLblsOverMax val=\"0\" />\r\n    </c:chart>\r\n    <c:spPr>\r\n        <a:solidFill>\r\n            <a:schemeClr val=\"bg1\" />\r\n        </a:solidFill>\r\n        <a:ln w=\"9525\" cap=\"flat\" cmpd=\"sng\" algn=\"ctr\">\r\n            <a:solidFill>\r\n                <a:schemeClr val=\"tx1\">\r\n                    <a:lumMod val=\"15000\" />\r\n                    <a:lumOff val=\"85000\" />\r\n                </a:schemeClr>\r\n            </a:solidFill>\r\n            <a:round />\r\n        </a:ln>\r\n        <a:effectLst />\r\n    </c:spPr>\r\n    <c:txPr>\r\n        <a:bodyPr />\r\n        <a:lstStyle />\r\n        <a:p>\r\n            <a:pPr>\r\n                <a:defRPr />\r\n            </a:pPr>\r\n            <a:endParaRPr lang=\"en-SE\" />\r\n        </a:p>\r\n    </c:txPr>\r\n    <c:printSettings>\r\n        <c:headerFooter />\r\n        <c:pageMargins b=\"0.75\" l=\"0.7\" r=\"0.7\" t=\"0.75\" header=\"0.3\" footer=\"0.3\" />\r\n        <c:pageSetup />\r\n    </c:printSettings>\r\n</c:chartSpace>";
            // Load chart1.xml
            XDocument chartDoc = XDocument.Parse(chart2);
            XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";

            // Extract categories and values
            var ser = chartDoc.Descendants(c + "ser").FirstOrDefault();

            var categories = ser
                .Descendants(c + "cat")
                .Descendants(c + "strCache")
                .Descendants(c + "pt")
                .Select(pt => pt.Element(c + "v")?.Value)
                .ToList();

            var values = ser
                .Descendants(c + "val")
                .Descendants(c + "numCache")
                .Descendants(c + "pt")
                .Select(pt => double.Parse(pt.Element(c + "v")?.Value ?? "0"))
                .ToList();

            // Dimensions for SVG
            int width = 600, height = 400;
            int margin = 50;
            int barWidth = 40;
            int spacing = 20;

            double maxValue = values.Max();
            double scaleY = (height - 2 * margin) / maxValue;

            // Create SVG root
            XNamespace svgNs = "http://www.w3.org/2000/svg";
            XElement svg = new XElement(svgNs + "svg",
                new XAttribute("xmlns", svgNs),
                new XAttribute("width", width),
                new XAttribute("height", height),
                new XElement(svgNs + "rect",  // Background
                    new XAttribute("x", 0),
                    new XAttribute("y", 0),
                    new XAttribute("width", width),
                    new XAttribute("height", height),
                    new XAttribute("fill", "white"))
            );

            // Draw bars
            for (int i = 0; i < values.Count; i++)
            {
                double barHeight = values[i] * scaleY;
                double x = margin + i * (barWidth + spacing);
                double y = height - margin - barHeight;

                svg.Add(new XElement(svgNs + "rect",
                    new XAttribute("x", x),
                    new XAttribute("y", y),
                    new XAttribute("width", barWidth),
                    new XAttribute("height", barHeight),
                    new XAttribute("fill", "steelblue")));

                // Add category label
                svg.Add(new XElement(svgNs + "text",
                    new XAttribute("x", x + barWidth / 2),
                    new XAttribute("y", height - margin + 15),
                    new XAttribute("text-anchor", "middle"),
                    new XAttribute("font-size", "12"),
                    categories[i]));

                // Add value label
                svg.Add(new XElement(svgNs + "text",
                    new XAttribute("x", x + barWidth / 2),
                    new XAttribute("y", y - 5),
                    new XAttribute("text-anchor", "middle"),
                    new XAttribute("font-size", "12"),
                    values[i].ToString()));
            }

            // Save to SVG file
            svg.Save("chart.svg");
        }
    }
}
