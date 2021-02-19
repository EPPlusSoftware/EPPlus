using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table.PivotTable;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Table.PivotTable
{
    [TestClass]
    public class PivotTableLargeStyleTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
        }
        [ClassCleanup]
        public static void Cleanup()
        {
        }
        [TestMethod]
        public void AddPivotAllStyle()
        {
            using (var p=OpenTemplatePackage("PivotStyleLarge.xlsx"))
            {
                var ws = p.Workbook.Worksheets[0];
                var pt = ws.PivotTables[0];

                var s0 = pt.Styles.AddButtonField(ePivotTableAxis.PageAxis, 2);
                s0.Style.Font.Color.SetColor(Color.Pink);

                var s1 =pt.Styles.AddButtonField(pt.Fields["FacilityName"]);
                s1.Style.Font.Color.SetColor(OfficeOpenXml.Drawing.eThemeSchemeColor.Accent1);

                var s2 = pt.Styles.AddLabel(pt.Fields["FacilityName"]);
                s2.Style.Font.Color.SetColor(Color.Green);

                var s3 = pt.Styles.AddButtonField(pt.Fields["SiteId"]);
                s3.Style.Font.Color.SetColor(Color.Blue);

                var s4 = pt.Styles.AddLabel(pt.Fields["SiteId"]);
                s4.Style.Font.Color.SetColor(Color.Cyan);
                s4.Conditions.Fields[0].Items.AddByValue(5D);
                s4.Conditions.Fields[0].Items.AddByValue(8D);
                s4.Conditions.Fields[0].Items.AddByValue(9D);

                var s5 = pt.Styles.AddData(pt.Fields["SiteId"], pt.Fields["ZipCode"], pt.Fields["Id"]);
                s5.Style.Fill.PatternType = ExcelFillStyle.DarkTrellis;
                s5.Style.Fill.BackgroundColor.SetColor(Color.Red);
                s5.Conditions.DataFields.Add(1);
                s5.Conditions.Fields[0].Items.AddByValue(1D);
                s5.Conditions.Fields[0].Items.AddByValue(2D);
                s5.Conditions.Fields[0].Items.AddByValue(3D);
                s5.Conditions.Fields[1].Items.AddByValue("02201");
                s5.Conditions.Fields[2].Items.AddByValue("1100");


                SaveWorkbook("PivotStyleLargeSaved.xlsx", p);
            }
        }
    }
}

