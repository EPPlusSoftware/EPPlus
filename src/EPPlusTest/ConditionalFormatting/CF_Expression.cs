using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.ConditionalFormatting.Rules;

namespace EPPlusTest.ConditionalFormatting
{
    [TestClass]
    public class CF_Expression : TestBase
    {
        [TestMethod]
        public void CustomExpressionShouldApply()
        {
            using (var pck = OpenPackage("CustomExpression.xlsx", true))
            {
                var sheet = pck.Workbook.Worksheets.Add("ExpressionSheet");

                var expression = sheet.Cells["A1:A10"].ConditionalFormatting.AddExpression();

                expression.Formula = "A1<5";
                sheet.Cells["A1:A10"].Formula = "ROW()";

                sheet.Cells["A1:A10"].Calculate();

                var expressionClass = (ExcelConditionalFormattingExpression)expression;

                Assert.IsTrue(expressionClass.ShouldApplyToCell(sheet.Cells["A1"]));
                Assert.IsFalse(expressionClass.ShouldApplyToCell(sheet.Cells["A6"]));
            }
        }
    }
}
