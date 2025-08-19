using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.Variables
{
    [TestClass]
    public class LetFunctionIntegrationTests : TestBase
    {
        [TestMethod]
        public void ShouldAddPrefixToParameters_WithCalculate()
        {
            using var package = OpenPackage("LetFunction_Calculate_ShouldAddParamPrefixes.xlsx", true);
            var sheet = package.Workbook.Worksheets.Add("LET params");
            sheet.Cells["A1"].Formula = "LET(x,A2, y, 2, x+y)";
            sheet.Cells["A2"].Value = 1;
            sheet.Calculate();
            package.Save();
        }

        [TestMethod]
        public void ShouldAddPrefixToParameters_WithoutCalculate()
        {
            using var package = OpenPackage("LetFunction_WithoutCalculate_ShouldAddParamPrefixes.xlsx", true);
            var sheet = package.Workbook.Worksheets.Add("LET params");
            sheet.Cells["A1"].Formula = "LET(x,A2, y, 2, x+y)";
            sheet.Cells["A2"].Value = 1;
            //sheet.Calculate();
            package.Save();
        }
    }
}
