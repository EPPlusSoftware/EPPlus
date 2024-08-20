using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Drawing;
using System.IO;
using OfficeOpenXml.ConditionalFormatting;
using System.Data;
namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableConditionalFormattingTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("PivotTableConditionalFormat.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Data1");
            LoadItemData(_ws);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            var dirName = _pck.File.DirectoryName;
            var fileName = _pck.File.FullName;
            SaveAndCleanup(_pck);
            File.Copy(fileName, dirName + "\\PivotTableConditionalFormatRead.xlsx", true);
        }
        internal ExcelPivotTable CreatePivotTable(ExcelWorksheet ws)
        {
            var pt = ws.PivotTables.Add(ws.Cells["A3"], _ws.Cells[_ws.Dimension.Address], "PivotTable1");
            pt.RowFields.Add(pt.Fields[0]);
            pt.ColumnFields.Add(pt.Fields[1]);
            pt.DataFields.Add(pt.Fields[3]);
            pt.DataFields.Add(pt.Fields[2]);
            pt.PageFields.Add(pt.Fields[4]);
            return pt;
        }
        [TestMethod]
        public void AddPivotCF_TowDataField()
        {
            var ws = _pck.Workbook.Worksheets.Add("FirstDataFieldGreaterThan");
            var pt = CreatePivotTable(ws);
            var rule = pt.ConditionalFormattings.Add(eExcelPivotTableConditionalFormattingRuleType.GreaterThan);            

            var area = rule.Areas.Add();
            area.Conditions.DataFields.Add(pt.DataFields[0]);
            rule.ConditionalFormatting.As.GreaterThan.Formula = "7.3";
            rule.ConditionalFormatting.Style.Font.Color.SetColor(Color.Red);

            var rule2 = pt.ConditionalFormattings.Add(eExcelPivotTableConditionalFormattingRuleType.GreaterThan);
            var area2 = rule2.Areas.Add();
            area2.Conditions.DataFields.Add(pt.DataFields[1]);
            rule2.ConditionalFormatting.As.GreaterThan.Formula = "50";
            rule2.ConditionalFormatting.Style.Font.Color.SetColor(Color.Blue);
        }
        [TestMethod]
        public void AddPivotCF_AddTwoPivotAreas()
        {
            var ws = _pck.Workbook.Worksheets.Add("FirstDataFieldGreaterThan");
            var pt = CreatePivotTable(ws);
            var rule = pt.ConditionalFormattings.Add(eExcelPivotTableConditionalFormattingRuleType.GreaterThan);

            //rule.Type = ePivotTableConditionalFormattingConditionType.Row; //row and column causes the workbook to become corrupt.
            rule.Scope = ePivotTableConditionalFormattingConditionScope.Field;
            var area = rule.Areas.Add();
            area.Conditions.DataFields.Add(pt.DataFields[0]);
            rule.ConditionalFormatting.As.GreaterThan.Formula = "7.3";
            rule.ConditionalFormatting.Style.Font.Color.SetColor(Color.Red);

            var area2 = rule.Areas.Add();
            area2.Conditions.DataFields.Add(pt.DataFields[1]);
            pt.Fields["Category"].Items.Refresh();
            area2.Conditions.Fields.Add(pt.Fields["Category"]);
            area2.Conditions.Fields[0].Items.Add(0);
        }
        [TestMethod]
        public void AddPivotCF_AddExtLstPivotFormattingData()
        {
            var ws = _pck.Workbook.Worksheets.Add("FirstDataFieldGreaterThan");
            var pt = CreatePivotTable(ws);
            var rule = pt.ConditionalFormattings.Add(eExcelPivotTableConditionalFormattingRuleType.DataBar);
            rule.Scope = ePivotTableConditionalFormattingConditionScope.Data;

            var area = rule.Areas.Add();
            area.Conditions.DataFields.Add(pt.DataFields[1]);

            var db = rule.ConditionalFormatting.As.DataBar;
        }
        [TestMethod]
        public void AddPivotCF_AddExtLstPivotFormatting()
        {
            var ws = _pck.Workbook.Worksheets.Add("FirstDataFieldBottomPercent");
            var pt = CreatePivotTable(ws);
            var rule = pt.ConditionalFormattings.Add(eExcelPivotTableConditionalFormattingRuleType.BottomPercent, pt.DataFields[0]);
            var area = rule.Areas[0];

            rule.Scope = ePivotTableConditionalFormattingConditionScope.Field;
            area.Conditions.Fields.Add(pt.Fields[0]);

            rule.Scope = ePivotTableConditionalFormattingConditionScope.Field;
            rule.Type = ePivotTableConditionalFormattingConditionType.Column;
            rule.ConditionalFormatting.As.TopBottom.Rank = 20;
            rule.ConditionalFormatting.Style.Fill.PatternType = ExcelFillStyle.Solid;
            rule.ConditionalFormatting.Style.Fill.BackgroundColor.SetColor(Color.Red);
        }
        [TestMethod]
        public void AddPivotCF_AddIconset()
        {
            var ws = _pck.Workbook.Worksheets.Add("FirstDataFieldBottomPercent");
            var pt = CreatePivotTable(ws);
            var rule = pt.ConditionalFormattings.Add(eExcelPivotTableConditionalFormattingRuleType.ThreeIconSet, pt.DataFields[0]);
            var area = rule.Areas[0];

            rule.Scope = ePivotTableConditionalFormattingConditionScope.Field;
            area.Conditions.Fields.Add(pt.Fields[0]);

            rule.ConditionalFormatting.As.ThreeIconSet.IconSet=eExcelconditionalFormatting3IconsSetType.Stars;
        }
    }
}

