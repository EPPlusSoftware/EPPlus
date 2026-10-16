/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/12/2026         EPPlus Software AB       Initial release, tests for issue #2572
 *************************************************************************************************/
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace EPPlusTest.Core.Worksheet
{
    /// <summary>
    /// Tests for the XML written by the worksheet XML writer.
    /// The prefix tests use a worksheet where the main schema has a namespace prefix (x:) and no default namespace,
    /// as written by the Open XML SDK. Any element written without the prefix ends up in no namespace.
    /// </summary>
    [TestClass]
    public class WorksheetXmlWriterTests : TestBase
    {
        private const string PrefixedWorkbook = "sc1083-sdk.xlsx";
        private const string CalcFlagsWorkbook = "sc1083-flags.xlsx";
        private const string WorksheetPart = "xl/worksheets/sheet1.xml";
        private const string SchemaMainX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

        /// <summary>
        /// Issue #2572: Shared formula children were written as &lt;f&gt; without the namespace prefix,
        /// so Excel treated the cells as constants.
        /// </summary>
        [TestMethod]
        public void SharedFormulaChildKeepsNamespacePrefix()
        {
            const string outputName = "WorksheetXmlWriter_SharedFormulaChild.xlsx";
            using (var p = OpenTemplatePackage(PrefixedWorkbook))
            {
                SaveWorkbook(outputName, p);
            }

            var doc = LoadPartXml(outputName, WorksheetPart);
            var nsm = CreateNamespaceManager(doc);

            //Guard: the shared formula children must be written, otherwise the test does not exercise the shared formula path.
            Assert.IsNotNull(doc.SelectSingleNode("//m:c[@r='C4']/m:f[@t='shared']", nsm), "C4 has no shared formula in the main namespace.");
            Assert.IsNotNull(doc.SelectSingleNode("//m:c[@r='C5']/m:f[@t='shared']", nsm), "C5 has no shared formula in the main namespace.");
            AssertNoElementsWithoutNamespace(doc);
        }

        /// <summary>
        /// Issue #2572: Icon set cfvo elements were written without the namespace prefix,
        /// so Excel reported the workbook as corrupt.
        /// </summary>
        [TestMethod]
        public void IconSetCfvoKeepsNamespacePrefix()
        {
            const string outputName = "WorksheetXmlWriter_IconSetCfvo.xlsx";
            using (var p = OpenTemplatePackage(PrefixedWorkbook))
            {
                var ws = p.Workbook.Worksheets[0];
                ws.ConditionalFormatting.AddThreeIconSet(ws.Cells["A2:A5"], eExcelconditionalFormatting3IconsSetType.Arrows);
                SaveWorkbook(outputName, p);
            }

            var doc = LoadPartXml(outputName, WorksheetPart);

            //Guard: the icon set must be written, otherwise the test does not exercise the cfvo path.
            Assert.IsTrue(doc.GetElementsByTagName("iconSet", ExcelPackage.schemaMain).Count > 0, "No iconSet element was written.");
            AssertNoElementsWithoutNamespace(doc);
        }

        /// <summary>
        /// Issue #2572: The children of x14:dxf in extLst conditional formatting were written without the main namespace prefix,
        /// so Excel silently dropped the format of the rule.
        /// </summary>
        [TestMethod]
        public void ExtLstDxfChildrenKeepNamespacePrefix()
        {
            const string outputName = "WorksheetXmlWriter_ExtLstDxf.xlsx";
            using (var p = OpenTemplatePackage(PrefixedWorkbook))
            {
                var ws = p.Workbook.Worksheets[0];
                var ws2 = p.Workbook.Worksheets.Add("Sheet2");
                ws2.Cells["A1"].Value = 3;

                //A reference to another worksheet forces the rule into extLst (x14).
                var cf = ws.ConditionalFormatting.AddExpression(ws.Cells["A2:A5"]);
                cf.Formula = "A2>Sheet2!$A$1";
                cf.Style.Font.Bold = true;
                cf.Style.Fill.BackgroundColor.Color = System.Drawing.Color.Yellow;
                SaveWorkbook(outputName, p);
            }

            var doc = LoadPartXml(outputName, WorksheetPart);

            //Guard: the rule must be written to extLst, otherwise the test does not exercise the dxf path.
            Assert.IsTrue(doc.GetElementsByTagName("conditionalFormatting", SchemaMainX14).Count > 0, "The rule was not written to extLst.");
            Assert.IsTrue(doc.GetElementsByTagName("dxf", SchemaMainX14).Count > 0, "No x14:dxf element was written.");
            AssertNoElementsWithoutNamespace(doc);
        }

        /// <summary>
        /// Issue #2572: The aca and ca attributes of an array formula leaked to the following formula cells,
        /// as the attribute string was only reset for cells with flags.
        /// </summary>
        [TestMethod]
        public void CalcAttributesDoNotLeakToNextFormulaCell()
        {
            const string outputName = "WorksheetXmlWriter_CalcAttributes.xlsx";
            using (var p = OpenTemplatePackage(CalcFlagsWorkbook))
            {
                SaveWorkbook(outputName, p);
            }

            var doc = LoadPartXml(outputName, WorksheetPart);
            var nsm = CreateNamespaceManager(doc);

            //Guard: the array formula must keep its calc attributes, otherwise there is nothing that can leak.
            var arrayFormula = doc.SelectSingleNode("//m:c[@r='D2']/m:f", nsm) as XmlElement;
            Assert.IsNotNull(arrayFormula, "D2 has no formula element.");
            Assert.AreEqual("1", arrayFormula.GetAttribute("aca"), "D2 lost its aca attribute on load or save.");

            foreach (var address in new[] { "C3", "C4", "C5" })
            {
                var formula = doc.SelectSingleNode($"//m:c[@r='{address}']/m:f", nsm) as XmlElement;
                Assert.IsNotNull(formula, $"{address} has no formula element.");
                Assert.IsFalse(formula.HasAttribute("aca"), $"{address} inherited aca from a previous cell.");
                Assert.IsFalse(formula.HasAttribute("ca"), $"{address} inherited ca from a previous cell.");
            }
        }

        private static XmlNamespaceManager CreateNamespaceManager(XmlDocument doc)
        {
            var nsm = new XmlNamespaceManager(doc.NameTable);
            nsm.AddNamespace("m", ExcelPackage.schemaMain);
            return nsm;
        }

        private static XmlDocument LoadPartXml(string fileName, string partName)
        {
            using (var fileStream = File.OpenRead(_worksheetPath + fileName))
            using (var zipStream = new OfficeOpenXml.Packaging.Ionic.Zip.ZipInputStream(fileStream))
            {
                OfficeOpenXml.Packaging.Ionic.Zip.ZipEntry entry;
                while ((entry = zipStream.GetNextEntry()) != null)
                {
                    if (entry.FileName.Equals(partName, StringComparison.OrdinalIgnoreCase))
                    {
                        var doc = new XmlDocument();
                        doc.Load(zipStream);
                        return doc;
                    }
                }
            }
            Assert.Fail($"Part {partName} is missing in {fileName}");
            return null;
        }

        private static void AssertNoElementsWithoutNamespace(XmlDocument doc)
        {
            var offenders = new List<string>();
            foreach (XmlElement element in doc.GetElementsByTagName("*"))
            {
                if (string.IsNullOrEmpty(element.NamespaceURI))
                {
                    var xml = element.OuterXml;
                    offenders.Add(xml.Length > 120 ? xml.Substring(0, 120) : xml);
                }
            }
            Assert.AreEqual(0, offenders.Count, "Elements without namespace: " + string.Join(" | ", offenders));
        }
    }
}