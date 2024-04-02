using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport.CssCollections;
using OfficeOpenXml.Export.HtmlExport;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Export.HtmlExport.Translators;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors;
using System.Drawing;

namespace EPPlusTest.Export.HtmlExport
{
    [TestClass]
    public class CssRuleCollectionTests
    {
        [TestMethod]
        public void ExportRangeWithNullBorderMergedCellsShouldNotThrow()
        {
            using (var package = new ExcelPackage())
            {

                var test = package.Workbook.Worksheets.Add("test");

                var wb = package.Workbook;

                test.Cells["A1"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Dotted;
                test.Cells["B1"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.DashDot;
                test.Cells["C1"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                test.Cells["C1"].Style.Fill.BackgroundColor.SetColor(Color.Red);

                var testStyle = new StyleXml(wb.Styles.CellXfs[0]);
                var style2 = new StyleXml(wb.Styles.CellXfs[1]);
                var testNull = test.Cells["A1"].Style.Styles.CellXfs[2];
                var style3 = new StyleXml(wb.Styles.CellXfs[2]);

                IBorder topLeft = testStyle.Border ?? null;
                IBorder bottom = style2.Border ?? null;
                IBorder right = style3.Border ?? null;

                var borderTranslator = new CssBorderTranslator(topLeft, bottom, right);
                var context = new TranslatorContext(new HtmlRangeExportSettings());

                var declarations = borderTranslator.GenerateDeclarationList(context);
            }
        }
    }
}
