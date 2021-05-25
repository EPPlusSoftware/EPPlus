using EPPlusTest;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ExternalReferences;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using System.Collections.Generic;
using System.IO;

namespace EPPlusTest.Core
{
    [TestClass]
    public class ExternalReferencesTest : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            //_pck = OpenPackage("ExternalReferences.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            //var dirName = _pck.File.DirectoryName;
            //var fileName = _pck.File.FullName;

            //SaveAndCleanup(_pck);

            //if (File.Exists(fileName)) File.Copy(fileName, dirName + "\\ExternalReferencesRead.xlsx", true);
            }
        [TestMethod]
        public void OpenAndReadExternalReferences()
        {
            var p = OpenTemplatePackage("ExtRef.xlsx");

            Assert.AreEqual(2, p.Workbook.ExternalReferences.Count);

            Assert.AreEqual(1D, p.Workbook.ExternalReferences[0].As.ExternalWorkbook.CachedWorksheets["sheet1"].CellValues["A2"].Value);
            Assert.AreEqual(12D, p.Workbook.ExternalReferences[0].As.ExternalWorkbook.CachedWorksheets["sheet1"].CellValues["C3"].Value);

            var c = 0;
            foreach(var cell in p.Workbook.ExternalReferences[0].As.ExternalWorkbook.CachedWorksheets["sheet1"].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(11, c);
        }

        [TestMethod]
        public void OpenAndCalculateExternalReferencesFromCache()
        {
            var p = OpenTemplatePackage("ExtRef.xlsx");

            p.Workbook.ClearFormulaValues();
            p.Workbook.Calculate();

            var ws = p.Workbook.Worksheets[0];
            Assert.AreEqual(2D, ws.Cells["E2"].Value);
            Assert.AreEqual(4D, ws.Cells["F2"].Value);
            Assert.AreEqual(6D, ws.Cells["G2"].Value);

            Assert.AreEqual(8D, ws.Cells["E3"].Value);
            Assert.AreEqual(16D, ws.Cells["F3"].Value);
            Assert.AreEqual(24D, ws.Cells["G3"].Value);

            Assert.AreEqual(20D, ws.Cells["H5"].Value);
            Assert.AreEqual(117D, ws.Cells["K5"].Value);

            Assert.AreEqual(111D, ws.Cells["H8"].Value);
            Assert.IsInstanceOfType(ws.Cells["J8"].Value, typeof(ExcelErrorValue));
            Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)ws.Cells["J8"].Value).Type);

            Assert.AreEqual(3D, ws.Cells["E10"].Value);
            Assert.IsInstanceOfType(ws.Cells["F10"].Value, typeof(ExcelErrorValue));
            Assert.AreEqual(eErrorType.Ref, ((ExcelErrorValue)ws.Cells["F10"].Value).Type);
        }
        [TestMethod]
        public void OpenAndCalculateExternalReferencesFromPackage()
        {
            var p = OpenTemplatePackage("ExtRef.xlsx");

            p.Workbook.ExternalReferences.Directories.Add(new DirectoryInfo(_testInputPathOptional));
            p.Workbook.ExternalReferences.LoadWorkbooks();
            p.Workbook.ExternalReferences[0].As.ExternalWorkbook.Package.Workbook.Calculate();
            p.Workbook.ClearFormulaValues();
            p.Workbook.Calculate();

            var ws = p.Workbook.Worksheets[0];
            Assert.AreEqual(3D, ws.Cells["D1"].Value);
            Assert.AreEqual(2D, ws.Cells["E2"].Value);
            Assert.AreEqual(2D, ws.Cells["E2"].Value);
            Assert.AreEqual(4D, ws.Cells["F2"].Value);
            Assert.AreEqual(6D, ws.Cells["G2"].Value);

            Assert.AreEqual(8D, ws.Cells["E3"].Value);
            Assert.AreEqual(16D, ws.Cells["F3"].Value);
            Assert.AreEqual(24D, ws.Cells["G3"].Value);

            //Assert.AreEqual(20D, ws.Cells["H5"].Value);
            Assert.AreEqual(117D, ws.Cells["K5"].Value);

            Assert.AreEqual(111D, ws.Cells["H8"].Value);
            Assert.AreEqual(20D, ws.Cells["J8"].Value);

            Assert.AreEqual(3D, ws.Cells["E10"].Value);
            Assert.AreEqual(19D, ws.Cells["F10"].Value);
        }

        [TestMethod]
        public void DeleteExternalReferences()
        {
            var p = OpenTemplatePackage("ExtRef.xlsx");

            Assert.AreEqual(2, p.Workbook.ExternalReferences.Count);

            p.Workbook.ExternalReferences.Delete(0);

            SaveAndCleanup(p);
        }

        [TestMethod]
        public void OpenAndReadExternalReferences1()
        {
            var p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

            Assert.AreEqual(62, p.Workbook.ExternalReferences.Count);

            var c = 0;
            foreach (var cell in p.Workbook.ExternalReferences[0].As.ExternalWorkbook.CachedWorksheets[0].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(104, c);
        }

        [TestMethod]
        public void DeleteExternalReferences1()
        {
            var p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

            p.Workbook.ExternalReferences.Delete(0);
            p.Workbook.ExternalReferences.Delete(8);
            p.Workbook.ExternalReferences.Delete(5);


            SaveAndCleanup(p);
        }

        [TestMethod]
        public void OpenAndReadExternalReferences2()
        {
            var p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

            Assert.AreEqual(204, p.Workbook.ExternalReferences.Count);

            var c = 0;
            foreach (var cell in p.Workbook.ExternalReferences[0].As.ExternalWorkbook.CachedWorksheets[0].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(104, c);
        }
        [TestMethod]
        public void OpenAndDeleteExternalReferences2()
        {
            var p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

            Assert.AreEqual(204, p.Workbook.ExternalReferences.Count);
            p.Workbook.ExternalReferences.Delete(103);
            Assert.AreEqual(203, p.Workbook.ExternalReferences.Count);
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void OpenAndCalculateExternalReferences1()
        {
            var p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

            p.Workbook.Calculate();
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void OpenAndCalculateExternalReferences2()
        {
            var p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

            Assert.AreEqual(204, p.Workbook.ExternalReferences.Count);
            p.Workbook.Calculate();
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void OpenAndClearExternalReferences1()
        {
            var p = OpenTemplatePackage("ExternalReferencesText1.xlsx");

            Assert.AreEqual(62, p.Workbook.ExternalReferences.Count);
            p.Workbook.ExternalReferences.Clear();
            Assert.AreEqual(0, p.Workbook.ExternalReferences.Count);
            SaveWorkbook("ExternalReferencesText1_Cleared.xlsx", p);
        }
        [TestMethod]
        public void OpenAndClearExternalReferences2()
        {
            var p = OpenTemplatePackage("ExternalReferencesText2.xlsx");

            Assert.AreEqual(204, p.Workbook.ExternalReferences.Count);
            p.Workbook.ExternalReferences.Clear();
            Assert.AreEqual(0, p.Workbook.ExternalReferences.Count);
            SaveWorkbook("ExternalReferencesText2_Cleared.xlsx", p);
        }

        [TestMethod]
        public void OpenAndClearExternalReferences3()
        {
            var p = OpenTemplatePackage("ExternalReferencesText3.xlsx");

            Assert.AreEqual(63, p.Workbook.ExternalReferences.Count);
            p.Workbook.ExternalReferences.Clear();
            Assert.AreEqual(0, p.Workbook.ExternalReferences.Count);
            SaveWorkbook("ExternalReferencesText3_Cleared.xlsx", p);
        }



        [TestMethod]
        public void OpenAndReadExternalReferences3()
        {
            var p = OpenTemplatePackage("ExternalReferencesText3.xlsx");

            Assert.AreEqual(63, p.Workbook.ExternalReferences.Count);

            var c = 0;
            foreach (var cell in p.Workbook.ExternalReferences[0].As.ExternalWorkbook.CachedWorksheets[0].CellValues)
            {
                Assert.IsNotNull(cell.Value);
                c++;
            }
            Assert.AreEqual(104, c);
        }
        [TestMethod]
        public void OpenAndCalculateExternalReferences3()
        {
            var p = OpenTemplatePackage("ExternalReferencesText3.xlsx");

            p.Workbook.Calculate();
            SaveAndCleanup(p);
        }
        [TestMethod]
        public void OpenAndReadExternalReferencesDdeOle()
        {
            var p = OpenTemplatePackage("dde.xlsx");

            Assert.AreEqual(6, p.Workbook.ExternalReferences.Count);

            Assert.AreEqual(eExternalLinkType.DdeLink, p.Workbook.ExternalReferences[0].ExternalLinkType);
            p.Workbook.ExternalReferences.Directories.Add(new DirectoryInfo("c:\\epplustest\\workbooks"));
            p.Workbook.ExternalReferences.LoadWorkbooks();

            var book3 = p.Workbook.ExternalReferences[3].As.ExternalWorkbook;
            Assert.AreEqual("c:\\epplustest\\workbooks\\fromwb1.xlsx", book3.File.FullName.ToLower());
            Assert.IsNotNull(book3.Package);
            var book4 = p.Workbook.ExternalReferences[4].As.ExternalWorkbook;
            Assert.AreEqual("c:\\epplustest\\testoutput\\extref.xlsx", book4.File.FullName.ToLower());
            Assert.IsNotNull(book4.Package);
            SaveAndCleanup(p);
        }

        [TestMethod]
        public void UpdateCacheShouldBeSameAsExcel()
        {
            var p = OpenTemplatePackage("ExtRef.xlsx");

            var er = p.Workbook.ExternalReferences[0].As.ExternalWorkbook;
            var excelCache = GetExternalCache(er);

            p.Workbook.ExternalReferences[0].As.ExternalWorkbook.UpdateCache();
            var epplusCache = GetExternalCache(er);

            foreach (var key in excelCache.Keys)
            {
                if(epplusCache.ContainsKey(key))
                {
                    //We remove any single quotes as excel adds seems to add ' to all worksheet names.
                    //We also remove any prefixing equal sign in teh GetExternalCache method.
                    Assert.AreEqual(excelCache[key].ToString().Replace("\'","").ToString(), epplusCache[key].ToString().Replace("\'",""));
                }
                else
                {
                    Assert.Fail($"Key:{key} are missing in EPPlus cache.");
                }
            }

            foreach (var key in epplusCache.Keys)
            {
                if (!excelCache.ContainsKey(key))
                {
                    Assert.Fail($"Key:{key} are missing in EPPlus cache.");
                }
            }

            SaveWorkbook("ExtRef_Updated.xlsx", p);
        }

        [TestMethod]
        public void AddExternalReferenceShouldBeSameAsExcel()
        {
            var p = OpenPackage("AddedExtRef.xlsx", true);
            var ws1=CreateWorksheet1(p);
            var ws2 = p.Workbook.Worksheets.Add("Sheet2");
            
            ws2.Cells["A1"].Value = 3;
            ws2.Names.Add("SheetDefinedName", ws2.Cells["A1"]);
            
            ws1.Cells["D2"].Formula = "Sheet2!SheetDefinedName";
            ws1.Cells["E2"].Formula = "Table1[[#This Row],[a]]+[1]Sheet1!$A2";
            ws1.Cells["F2"].Formula = "Table1[[#This Row],[b]]+[1]Sheet1!$B2";
            ws1.Cells["G2"].Formula = "Table1[[#This Row],[c]]+[1]Sheet1!$C2";
            ws1.Cells["E3"].Formula = "Table1[[#This Row],[a]]+[1]Sheet1!$A3";
            ws1.Cells["F3"].Formula = "Table1[[#This Row],[b]]+[1]Sheet1!$B3";
            ws1.Cells["G3"].Formula = "Table1[[#This Row],[c]]+'[1]Sheet1'!$C3";
            var er = p.Workbook.ExternalReferences.AddWorkbook(new FileInfo(_testInputPath + "externalreferences\\FromWB1.xlsx"));
            
            ws1.Cells["G5"].Formula = $"[{er.Index}]Sheet1!FromF2*[{er.Index}]!CellH5";

            er.UpdateCache();
            ws1.Calculate();


            Assert.AreEqual(2220D, ws1.Cells["G5"].Value);
            SaveAndCleanup(p);
        }

        private static ExcelWorksheet CreateWorksheet1(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            ws.SetValue(1, 1, "a");
            ws.SetValue(1, 2, "b");
            ws.SetValue(1, 3, "c");
            ws.SetValue(2, 1, 1D);
            ws.SetValue(2, 2, 2D);
            ws.SetValue(2, 3, 3D);
            ws.SetValue(3, 1, 4D);
            ws.SetValue(3, 2, 8D);
            ws.SetValue(3, 3, 12D);

            var tbl = ws.Tables.Add(ws.Cells["A1:C3"], "Table1");
            //Create Table
            tbl.TableStyle = OfficeOpenXml.Table.TableStyles.Medium2;
            return ws;
        }

        private Dictionary<string, object> GetExternalCache(ExcelExternalWorkbook ewb)
        {
            var d=new Dictionary<string, object>();
            foreach(ExcelExternalWorksheet ws in ewb.CachedWorksheets)
            {
                foreach(ExcelExternalCellValue v in ws.CellValues)
                {
                    d.Add(ws.Name + v.Address, v.Value);
                }
                foreach (ExcelExternalDefinedName n in ws.CachedNames)
                {
                    if (n.RefersTo.StartsWith("="))
                    {
                        d.Add(ws.Name + n.Name, n.RefersTo.Substring(1));
                    }
                    else
                    {
                        d.Add(ws.Name + n.Name, n.RefersTo);
                    }
                }
            }
            foreach (ExcelExternalDefinedName n in ewb.CachedNames)
            {
                if (n.RefersTo.StartsWith("="))
                {
                    d.Add(n.Name, n.RefersTo.Substring(1));
                }
                else
                {
                    d.Add(n.Name, n.RefersTo);
                }
            }
            return d;
        }
    }
}
