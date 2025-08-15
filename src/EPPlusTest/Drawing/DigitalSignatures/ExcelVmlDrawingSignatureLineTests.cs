using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Xml;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Vml;
using OfficeOpenXml.Utils.Security;

namespace EPPlusTest.Drawing.DigitalSignatures
{
    [TestClass]
    public class ExcelVmlDrawingSignatureLineTests : TestBase
    {
        [TestMethod]
        public void CreateVmlSignatureLine()
        {
            using (ExcelPackage package = OpenPackage("UnsignedSignatureLine.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("SigLine");
                var node = ws.VmlDrawings.AddDigitalSignatureLineDrawing(SecurityUtil.CreateSecureGuid());

                Assert.AreEqual("_x0000_s1025", node.Attributes.GetNamedItem("id").Value);
                Assert.AreEqual("#_x0000_t75", node.Attributes.GetNamedItem("type").Value);

                node = node.SelectSingleNode("//v:shape", ws.VmlDrawings.NameSpaceManager);

                var imageData = node.SelectSingleNode("//v:imagedata", ws.VmlDrawings.NameSpaceManager);
                Assert.AreEqual("rId1", imageData.Attributes.GetNamedItem("relid").Value);

                var lockShapeEl = ws.VmlDrawings.VmlDrawingXml.CreateElement("o", "lock", ExcelPackage.schemaMicrosoftOffice);
                lockShapeEl.SetAttribute("ext", ExcelPackage.schemaMicrosoftVml, "edit");
                lockShapeEl.SetAttribute("ungrouping", "t");
                lockShapeEl.SetAttribute("rotation", "t");
                lockShapeEl.SetAttribute("cropping", "t");
                lockShapeEl.SetAttribute("verticies", "t");
                lockShapeEl.SetAttribute("text", "t");
                lockShapeEl.SetAttribute("grouping", "t");

                var lockNode = node.SelectSingleNode("o:lock", ws.VmlDrawings.NameSpaceManager);

                for(int i = 0; i < lockShapeEl.Attributes.Count; i++)
                {
                    var expectedAttribute = lockShapeEl.Attributes[i];
                    var actualAttribute = lockNode.Attributes[i];
                    Assert.AreEqual(expectedAttribute.Value, actualAttribute.Value);
                }

                var sigLineNode = node.SelectSingleNode("o:signatureline", ws.VmlDrawings.NameSpaceManager);
                var provId = sigLineNode.Attributes.GetNamedItem("provid");
                Assert.AreEqual("{00000000-0000-0000-0000-000000000000}", provId.Value);
                Assert.AreEqual("t", sigLineNode.Attributes.GetNamedItem("issignatureline").Value);
                var anchor = node.SelectSingleNode("//x:Anchor", ws.VmlDrawings.NameSpaceManager);
                Assert.AreEqual("0, 0, 0, 0, 4, 0, 6, 8", anchor.InnerText);
            }
        }

        [TestMethod]
        public void UpdateShapeTypeNode()
        {
            using (ExcelPackage p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("ShapeTypes");
                var node = ws.VmlDrawings.UpdateShapeTypeForSignatureLine();

                Assert.AreEqual("_x0000_t75", node.Attributes.GetNamedItem("id").Value);
                Assert.AreEqual("21600,21600", node.Attributes.GetNamedItem("coordsize").Value);
                Assert.AreEqual("75", node.Attributes.GetNamedItem("spt").Value);
                Assert.AreEqual("t", node.Attributes.GetNamedItem("preferrelative").Value);
                Assert.AreEqual("m@4@5l@4@11@9@11@9@5xe", node.Attributes.GetNamedItem("path").Value);
                Assert.AreEqual("f", node.Attributes.GetNamedItem("filled").Value);
                Assert.AreEqual("f", node.Attributes.GetNamedItem("stroked").Value);

                Assert.AreEqual("miter", node.ChildNodes[0].Attributes.GetNamedItem("joinstyle").Value);

                var tempElement = ws.VmlDrawings.VmlDrawingXml.CreateNode(XmlNodeType.Element, "anElement", node.NamespaceURI);
                ExcelVmlDrawingSignatureLine.CreateFormulaElementAsChildOf(tempElement);

                Assert.AreEqual(tempElement.InnerXml, node.ChildNodes[1].OuterXml);

                Assert.AreEqual("f", node.ChildNodes[2].Attributes.GetNamedItem("extrusionok").Value);
                Assert.AreEqual("t", node.ChildNodes[2].Attributes.GetNamedItem("gradientshapeok").Value);
                Assert.AreEqual("rect", node.ChildNodes[2].Attributes.GetNamedItem("connecttype").Value);

                Assert.AreEqual("edit", node.ChildNodes[3].Attributes.GetNamedItem("ext").Value);
                Assert.AreEqual("t", node.ChildNodes[3].Attributes.GetNamedItem("aspectratio").Value);
            }
        }

        [TestMethod]
        public void CreateSignatureLineStamp()
        {
            using (ExcelPackage package = OpenPackage("UnsignedSignatureLineStamp.xlsx", true))
            {
                var ws = package.Workbook.Worksheets.Add("SignatureLineStamp");
                var stamp = ws.VmlDrawings.AddSignatureLineStamp();

                var provId = stamp.TopNode.SelectSingleNode("//o:signatureline", ws.VmlDrawings.NameSpaceManager).Attributes.GetNamedItem("provid");
                Assert.AreEqual("{000CD6A4-0000-0000-C000-000000000046}", provId.Value);
                var anchor = stamp.TopNode.SelectSingleNode("//x:Anchor", ws.VmlDrawings.NameSpaceManager);
                Assert.AreEqual("0, 0, 0, 0, 2, 0, 8, 0", anchor.InnerXml);
            }
        }
    }
}
