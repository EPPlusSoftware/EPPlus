using System;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using EPPlusTest.Properties;
using OfficeOpenXml.FormulaParsing;


namespace EPPlusTest
{
  
  
    

    [TestClass]
    public class AmanaIssues 
    {
        [TestMethod]
        public void SUMMIF_Formula_Issue()
        {

            //Issue: SUMMIF can't be calculated correctly, if row or column number is out of the range

            var excelTestFile = Resources.TestDoc_SharedFormula_xlsx;

            using (MemoryStream excelStream = new MemoryStream())

            {

                excelStream.Write(excelTestFile, 0, excelTestFile.Length);

                using (ExcelPackage exlPackage = new ExcelPackage(excelStream))


                {
                    exlPackage.Workbook.Calculate();

                    var value1 = exlPackage.Workbook.Worksheets[1].Cells["J10"].Value;

                    var value2 = exlPackage.Workbook.Worksheets[1].Cells["J11"].Value;

                    Assert.IsTrue(value1.Equals(1.95583D));

                    Assert.IsTrue(value2.Equals(7.84515D));

                }

            }

        }
    }
}