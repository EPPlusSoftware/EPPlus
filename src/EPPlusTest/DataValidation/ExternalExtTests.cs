using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.Sparkline;
using System.IO;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Slicer;

namespace EPPlusTest.DataValidation
{
    [TestClass]
    public class ExternalExtTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            //_pck = OpenPackage("ExternalReferences.xlsx", true);
            var outDir = _worksheetPath + "ExternalDataValidations";
            if (!Directory.Exists(outDir)) Directory.CreateDirectory(outDir);
        }

        //Ensures no save or load errors
        internal void SaveAndLoadAndSave(in ExcelPackage pck)
        {
            var file = pck.File;

            var stream = new MemoryStream();
            pck.SaveAs(stream);

            var loadedPackage = new ExcelPackage(stream);

            loadedPackage.File = file;

            SaveAndCleanup(loadedPackage);
        }

        internal void AddDataValidations(ref ExcelWorksheet ws, bool isExtLst = false, string extSheetName = "", bool many = false)
        {
            if(isExtLst)
            {
                var intValidation = ws.DataValidations.AddIntegerValidation("A1");
                intValidation.Operator = ExcelDataValidationOperator.equal;
                intValidation.Formula.ExcelFormula = extSheetName + "!A1";
            }
            else
            {
                var intValidation = ws.DataValidations.AddIntegerValidation("A1");
                intValidation.Formula.Value = 1;
                intValidation.Formula2.Value = 3;
            }

            if(many)
            {
                var timeValidation = ws.DataValidations.AddTimeValidation("B1");
                timeValidation.Operator = ExcelDataValidationOperator.between;

                if (isExtLst)
                {
                    timeValidation.Formula.ExcelFormula = extSheetName + "!B1";
                    timeValidation.Formula2.ExcelFormula = extSheetName + "!B2";


                }
                else
                {
                    timeValidation.Formula.ExcelFormula = "B1";
                    timeValidation.Formula.ExcelFormula = "B2";
                }
            }
        }

        [TestMethod]
        public void LocalDataValidationsShouldWorkWithExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\LocalDVExternalCF.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

                AddDataValidations(ref ws, false);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void LocalDataValidationsShouldWorkWithManyExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\LocalDVManyExternalCF.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 2, 2, 2), Color.Red);

                AddDataValidations(ref ws, false);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ManyLocalDataValidationsShouldWorkWithExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\ManyLocalDVExternalCF.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 2, 2, 2), Color.Red);

                AddDataValidations(ref ws, false, "", true);
                SaveAndLoadAndSave(pck);
            }

        }

        [TestMethod]
        public void ManyLocalDataValidationsShouldWorkWithManyExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\ManyLocalDVManyExternalCF.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);


                AddDataValidations(ref ws, false, "", true);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void LocalDataValidationsShouldWorkWithManyExtLstSparklines()
        {
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.SparklineGroups.Add(eSparklineType.Line, new ExcelAddress(1, 1, 5, 1), new ExcelAddress(1, 2, 5, 2));
                ws.SparklineGroups.Add(eSparklineType.Line, new ExcelAddress(1, 3, 5, 3), new ExcelAddress(1, 4, 5, 4));

                AddDataValidations(ref ws, false);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ExtDataValidationsShouldWorkWithExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\ExtDVExternalCF.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

                AddDataValidations(ref ws, true, extSheet.Name);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ManyExtDataValidationsShouldWorkWithExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\ManyExtDVExternalCF.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

                AddDataValidations(ref ws, true, extSheet.Name, true);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ManyExtDataValidationsShouldWorkWithManyExtLstConditionalFormattings()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\ManyExtDVManyExternalCF.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);
                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 2, 2, 2), Color.Red);

                AddDataValidations(ref ws, true, extSheet.Name, true);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void ExtDataValidationsShouldWorkWithAllOtherExts()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\ExtDVAllOtherExts.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["A1:A5"], ws.Cells["B1:B5"]);
                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

                ExcelRange range = ws.Cells[1, 1, 4, 1];
                ExcelTable table = ws.Tables.Add(range, "TestTable");
                table.StyleName = "None";

                ExcelTableSlicer slicer = ws.Drawings.AddTableSlicer(table.Columns[0]);

                AddDataValidations(ref ws, true, extSheet.Name, true);
                SaveAndLoadAndSave(pck);
            }
        }

        [TestMethod]
        public void LocalDataValidationsShouldWorkWithAllOtherExts()
        {
            using (var pck = OpenPackage("ExternalDataValidations\\LocalDVAllOtherExts.xlsx", true))
            {
                var ws = pck.Workbook.Worksheets.Add("conditionalFormattingsTest");
                var extSheet = pck.Workbook.Worksheets.Add("extAddressSheet");

                ws.SparklineGroups.Add(eSparklineType.Line, ws.Cells["A1:A5"], ws.Cells["B1:B5"]);
                ws.ConditionalFormatting.AddDatabar(new ExcelAddress(1, 1, 2, 1), Color.Blue);

                ExcelRange range = ws.Cells[1, 1, 4, 1];
                ExcelTable table = ws.Tables.Add(range, "TestTable");
                table.StyleName = "None";

                ExcelTableSlicer slicer = ws.Drawings.AddTableSlicer(table.Columns[0]);

                AddDataValidations(ref ws, false);
                SaveAndLoadAndSave(pck);
            }
        }
    }
}
