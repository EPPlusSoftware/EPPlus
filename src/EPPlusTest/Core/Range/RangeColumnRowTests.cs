﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Drawing;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeColumnRowTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("Range_RowColumn.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void Column_SetWidthBestFitAndStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Width");
            ws.Cells["A1:E5"].EntireColumn.Width = 30;
            ws.Cells["A1:E5"].EntireColumn.Style.Fill.SetBackground(Color.Red);

            ws.Cells["C10:C20"].EntireColumn.BestFit=true;

            Assert.AreEqual(30, ws.Cells["A1"].EntireColumn.Width);
            Assert.AreEqual(30, ws.Cells["C1"].EntireColumn.Width);
            Assert.AreEqual(30, ws.Cells["D1"].EntireColumn.Width);
            Assert.IsFalse(ws.Cells["B1"].EntireColumn.BestFit);
            Assert.IsTrue(ws.Cells["C1"].EntireColumn.BestFit);
            Assert.IsFalse(ws.Cells["D1"].EntireColumn.BestFit);

            Assert.AreEqual("FFFF0000", ws.Cells["E100"].EntireColumn.Style.Fill.BackgroundColor.Rgb);
        }

        [TestMethod]
        public void Column_SetPhonetic()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Phonetic");
            ws.Cells["D1:G5"].EntireColumn.Phonetic = true;
            ws.Cells["E1000:E5000"].EntireColumn.Phonetic = false;

            Assert.IsFalse(ws.Cells["C1"].EntireColumn.Phonetic);
            Assert.IsTrue(ws.Cells["D1"].EntireColumn.Phonetic);
            Assert.IsFalse(ws.Cells["E1"].EntireColumn.Phonetic);
            Assert.IsTrue(ws.Cells["G1"].EntireColumn.Phonetic);
            Assert.IsFalse(ws.Cells["H1"].EntireColumn.Phonetic);
        }

        [TestMethod]
        public void Column_SetHidden()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Hidden");
            ws.Cells["F1:J5"].EntireColumn.Hidden = true;
            ws.Cells["G1"].EntireColumn.Hidden = false;

            Assert.IsFalse(ws.Cells["E10"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["G10"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["K10"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["F1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["H1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["I1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["J1"].EntireColumn.Hidden);
        }
        [TestMethod]
        public void Column_CollapsChildren_Right_L()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Right_HideL");
            SetupColumnOutlineRight(ws);
            ws.Cells["L1"].EntireColumn.CollapseChildren(false);

            Assert.IsTrue(ws.Cells["A1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["K1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["L1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["M1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["N1"].EntireColumn.Hidden);
        }
        [TestMethod]
        public void Column_CollapsChildren_Right_P()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Right_HideP");
            SetupColumnOutlineRight(ws);
            ws.Cells["P:P"].EntireColumn.CollapseChildren(true);

            Assert.IsFalse(ws.Cells["P1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["M1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["N1"].EntireColumn.Hidden);

            Assert.IsFalse(ws.Cells["A1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["K1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["L1"].EntireColumn.Hidden);
        }

        [TestMethod]
        public void Column_CollapsChildren_Right_D()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Right_HideD");
            SetupColumnOutlineRight(ws);
            ws.Cells["D:D"].EntireColumn.CollapseChildren(true);

            Assert.IsFalse(ws.Cells["P1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["M1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["N1"].EntireColumn.Hidden);

            Assert.IsTrue(ws.Cells["A1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["C1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["D1"].EntireColumn.Hidden);
        }
        [TestMethod]
        public void Column_CollapsChildren_SetLevel1()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Right_Level1");
            SetupColumnOutlineRight(ws);
            ws.Cells["A:P"].EntireColumn.SetVisibleOutlineLevel(1);

            Assert.IsTrue(ws.Cells["M1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["C1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["B1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["A1"].EntireColumn.Hidden);

            Assert.IsFalse(ws.Cells["D1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["N1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["K1"].EntireColumn.Hidden);
        }
        [TestMethod]
        public void Column_CollapsChildren_SetLevel1_ExpandedChildren()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Right_Level1_ec");
            SetupColumnOutlineRight(ws);
            ws.Cells["A:P"].EntireColumn.SetVisibleOutlineLevel(1, false);

            Assert.IsTrue(ws.Cells["M1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["C1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["B1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["A1"].EntireColumn.Hidden);

            Assert.IsFalse(ws.Cells["D1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["N1"].EntireColumn.Hidden);
            Assert.IsFalse(ws.Cells["K1"].EntireColumn.Hidden);
        }
        [TestMethod]
        public void Column_CollapsChildren_SetLevel1_AlreadyCollapsed()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Right_Level1_ac");
            SetupColumnOutlineRight(ws);
            ws.Cells["A:P"].EntireColumn.CollapseChildren(true);
            ws.Cells["A:P"].EntireColumn.SetVisibleOutlineLevel(2, false);

            //Assert.IsTrue(ws.Cells["M1"].EntireColumn.Hidden);
            //Assert.IsTrue(ws.Cells["C1"].EntireColumn.Hidden);
            //Assert.IsTrue(ws.Cells["B1"].EntireColumn.Hidden);
            //Assert.IsTrue(ws.Cells["A1"].EntireColumn.Hidden);

            //Assert.IsFalse(ws.Cells["D1"].EntireColumn.Hidden);
            //Assert.IsFalse(ws.Cells["N1"].EntireColumn.Hidden);
            //Assert.IsFalse(ws.Cells["K1"].EntireColumn.Hidden);
        }

        [TestMethod]


        public void Column_CollapsChildren_Left()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Level0_Left");
            SetupColumnOutlineLeft(ws);

            ws.Cells["B1:C1"].EntireColumn.CollapseChildren(false);
            ws.Cells["N1"].EntireColumn.CollapseChildren(true);
            ws.Cells["J:L"].EntireColumn.CollapseChildren(true);

            Assert.IsFalse(ws.Cells["C1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["D1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["E10"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["G10"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["F1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["H1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["I1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["J1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["K10"].EntireColumn.Hidden);

            Assert.IsTrue(ws.Cells["C3"].EntireColumn.Collapsed);
            Assert.IsTrue(ws.Cells["N3"].EntireColumn.Collapsed);
            Assert.IsTrue(ws.Cells["K3"].EntireColumn.Collapsed);
            Assert.IsTrue(ws.Cells["L3"].EntireColumn.Collapsed);
        }
        [TestMethod]
        public void Column_CollapsChildren_Left_A()
        {
            var ws = _pck.Workbook.Worksheets.Add("Column_Collapsed_Level0_Left_A");
            SetupColumnOutlineLeft(ws);

            ws.Cells["A1"].EntireColumn.CollapseChildren(false);

            Assert.IsFalse(ws.Cells["A1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["C1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["D1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["E10"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["G10"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["F1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["H1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["I1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["J1"].EntireColumn.Hidden);
            Assert.IsTrue(ws.Cells["P10"].EntireColumn.Hidden);
        }

        [TestMethod]
        public void Row_CollapsChildren_Top()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_Collapsed_Level0");
            SetupRowOutlineTop(ws);

            ws.Cells["A3:A4"].EntireRow.CollapseChildren(false);
            ws.Cells["A10"].EntireRow.CollapseChildren(true);
            ws.Cells["13:13"].EntireRow.CollapseChildren(true);

            Assert.IsFalse(ws.Cells["C1"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A4"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A5"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A6"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A14"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["H15"].EntireRow.Hidden);

            Assert.IsTrue(ws.Cells["A3"].EntireRow.Collapsed);
            Assert.IsTrue(ws.Cells["A10"].EntireRow.Collapsed);
            Assert.IsTrue(ws.Cells["A13"].EntireRow.Collapsed);
        }

        [TestMethod]
        public void Row_CollapsChildren_TopSummaryTop()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_Collapsed_Level0_Below");
            SetupRowOutlineBelow(ws);
            ws.Cells["A13"].EntireRow.CollapseChildren(false);
            ws.Cells["A2"].EntireRow.CollapseChildren(false);

            Assert.IsTrue(ws.Cells["A1"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A2"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A12"].EntireRow.Hidden);
            Assert.IsFalse(ws.Cells["A13"].EntireRow.Hidden);

            Assert.IsTrue(ws.Cells["A2"].EntireRow.Collapsed);
            Assert.IsTrue(ws.Cells["A13"].EntireRow.Collapsed);
        }
        [TestMethod]
        public void Row_ExpandToLevel3()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_Collapsed_ExpandToLevel3");
            SetupRowOutlineBelow(ws);

            ws.Cells["A1:A19"].EntireRow.SetVisibleOutlineLevel(1);

            Assert.IsTrue(ws.Cells["A1"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A3"].EntireRow.Hidden);

            Assert.IsFalse(ws.Cells["A4"].EntireRow.Hidden);
            Assert.IsFalse(ws.Cells["A13"].EntireRow.Hidden);

            Assert.IsTrue(ws.Cells["A14"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["A16"].EntireRow.Hidden);

            Assert.IsFalse(ws.Cells["A17"].EntireRow.Hidden);
            Assert.IsFalse(ws.Cells["A19"].EntireRow.Hidden);
        }

        [TestMethod]
        public void Column_SetStyleName()
        {
            var styleName = "Green Fill";
            var ns = _pck.Workbook.Styles.CreateNamedStyle(styleName);
            ns.Style.Fill.SetBackground(Color.Green);
            var ws = _pck.Workbook.Worksheets.Add("Column_StyleName"); 
            
            ws.Cells["C15:J20"].EntireColumn.StyleName = "Green Fill";

            Assert.AreEqual("Green Fill", ws.Cells["E10"].EntireColumn.StyleName);
        }
        [TestMethod]
        public void Row_SetStyle()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_Style");

            ws.Cells["C15:J20"].EntireRow.Style.Font.Color.SetAuto();
            ws.Cells["C15:J20"].EntireRow.Style.Font.Bold = true; ;

            Assert.IsTrue(ws.Cells["E15"].EntireRow.Style.Font.Color.Auto);
            Assert.IsTrue(ws.Cells["E15"].EntireRow.Style.Font.Bold);
            Assert.IsTrue(ws.Cells["E20"].EntireRow.Style.Font.Color.Auto);
            Assert.IsTrue(ws.Cells["E20"].EntireRow.Style.Font.Bold);

            Assert.IsFalse(ws.Cells["E21"].EntireRow.Style.Font.Color.Auto);
            Assert.IsFalse(ws.Cells["E14"].EntireRow.Style.Font.Color.Auto);
        }

        [TestMethod]
        public void Row_SetStyleName()
        {
            var styleName = "Blue Fill";
            var ns = _pck.Workbook.Styles.CreateNamedStyle(styleName);
            ns.Style.Fill.SetBackground(Color.Blue);
            var ws = _pck.Workbook.Worksheets.Add("Row_StyleName");

            ws.Cells["C15:J20"].EntireRow.StyleName = styleName;

            Assert.AreEqual("Blue Fill", ws.Cells["E16"].EntireRow.StyleName);
        }
        [TestMethod]
        public void Row_SetPhonetic()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_Phonetic");

            ws.Cells["G15:K20"].EntireRow.Phonetic = true;
            ws.Cells["I17:J18"].EntireRow.Phonetic = false;

            Assert.IsFalse(ws.Cells["F14"].EntireRow.Phonetic);
            Assert.IsFalse(ws.Cells["I17"].EntireRow.Phonetic);
            Assert.IsFalse(ws.Cells["J18"].EntireRow.Phonetic);
            Assert.IsFalse(ws.Cells["L21"].EntireRow.Phonetic);

            Assert.IsTrue(ws.Cells["G15"].EntireRow.Phonetic);
            Assert.IsTrue(ws.Cells["H16"].EntireRow.Phonetic);
            Assert.IsTrue(ws.Cells["K19"].EntireRow.Phonetic);
        }
        [TestMethod]
        public void Row_SetHidden()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_Hidden");

            ws.Cells["G15:K20"].EntireRow.Hidden = true;
            ws.Cells["I17:J18"].EntireRow.Hidden = false;

            Assert.IsFalse(ws.Cells["F14"].EntireRow.Hidden);
            Assert.IsFalse(ws.Cells["I17"].EntireRow.Hidden);
            Assert.IsFalse(ws.Cells["J18"].EntireRow.Hidden);
            Assert.IsFalse(ws.Cells["L21"].EntireRow.Hidden);

            Assert.IsTrue(ws.Cells["G15"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["H16"].EntireRow.Hidden);
            Assert.IsTrue(ws.Cells["K19"].EntireRow.Hidden);
        }
        [TestMethod]
        public void Row_SetCollapsed()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_Collapsed");

            ws.Cells["G15:K20"].EntireRow.Collapsed = true;
            ws.Cells["I17:J18"].EntireRow.Collapsed = false;

            Assert.IsFalse(ws.Cells["F14"].EntireRow.Collapsed);
            Assert.IsFalse(ws.Cells["I17"].EntireRow.Collapsed);
            Assert.IsFalse(ws.Cells["J18"].EntireRow.Collapsed);
            Assert.IsFalse(ws.Cells["L21"].EntireRow.Collapsed);

            Assert.IsTrue(ws.Cells["G15"].EntireRow.Collapsed);
            Assert.IsTrue(ws.Cells["H16"].EntireRow.Collapsed);
            Assert.IsTrue(ws.Cells["K19"].EntireRow.Collapsed);
        }
        [TestMethod]
        public void Row_SetOutlineLevel()
        {
            var ws = _pck.Workbook.Worksheets.Add("Row_OutlineLevel");

            ws.Cells["G15:K20"].EntireRow.OutlineLevel = 1;
            ws.Cells["I17:J18"].EntireRow.OutlineLevel = 2;
            Assert.AreEqual(0, ws.Cells["F14"].EntireRow.OutlineLevel);
            Assert.AreEqual(2, ws.Cells["I17"].EntireRow.OutlineLevel);
            Assert.AreEqual(2, ws.Cells["J18"].EntireRow.OutlineLevel);
            Assert.AreEqual(0, ws.Cells["L21"].EntireRow.OutlineLevel);

            Assert.AreEqual(1, ws.Cells["G15"].EntireRow.OutlineLevel);
            Assert.AreEqual(1, ws.Cells["H16"].EntireRow.OutlineLevel);
            Assert.AreEqual(1, ws.Cells["K19"].EntireRow.OutlineLevel);
        }

        private static void SetupColumnOutlineRight(ExcelWorksheet ws)
        {
            ws.OutLineSummaryRight = true;
            ws.Cells["A1"].EntireColumn.OutlineLevel = 3;
            ws.Cells["B:B"].EntireColumn.OutlineLevel = 2;
            ws.Cells["C:C"].EntireColumn.OutlineLevel = 2;
            ws.Cells["D:J"].EntireColumn.OutlineLevel = 1;
            ws.Cells["K:K"].EntireColumn.OutlineLevel = 1;
            //ws.Cells["L:L"].EntireColumn.OutlineLevel = 0;
            ws.Cells["M:M"].EntireColumn.OutlineLevel = 2;
            ws.Cells["N:O"].EntireColumn.OutlineLevel = 1;
        }

        private static void SetupColumnOutlineLeft(ExcelWorksheet ws)
        {
            ws.OutLineSummaryRight = false;
            ws.Cells["A1"].EntireColumn.OutlineLevel = 1;
            ws.Cells["B:B"].EntireColumn.OutlineLevel = 2;
            ws.Cells["C:C"].EntireColumn.OutlineLevel = 2;
            ws.Cells["D:J"].EntireColumn.OutlineLevel = 3;
            ws.Cells["K:K"].EntireColumn.OutlineLevel = 3;
            ws.Cells["L:M"].EntireColumn.OutlineLevel = 4;
            ws.Cells["N:N"].EntireColumn.OutlineLevel = 3;
            ws.Cells["O:P"].EntireColumn.OutlineLevel = 4;
        }
        private static void SetupRowOutlineTop(ExcelWorksheet ws)
        {
            ws.OutLineSummaryBelow = false;
            ws.Cells["A1"].EntireRow.OutlineLevel = 1;
            ws.Cells["A2"].EntireRow.OutlineLevel = 2;
            ws.Cells["A3"].EntireRow.OutlineLevel = 2;
            ws.Cells["A4:B9"].EntireRow.OutlineLevel = 3;
            ws.Cells["C10"].EntireRow.OutlineLevel = 3;
            ws.Cells["A11:A12"].EntireRow.OutlineLevel = 4;
            ws.Cells["A13:A13"].EntireRow.OutlineLevel = 3;
            ws.Cells["B14:B15"].EntireRow.OutlineLevel = 4;
        }

        private static void SetupRowOutlineBelow(ExcelWorksheet ws)
        {
            ws.OutLineSummaryBelow = true;
            ws.Cells["A1"].EntireRow.OutlineLevel = 3;
            ws.Cells["A2"].EntireRow.OutlineLevel = 2;
            ws.Cells["A3"].EntireRow.OutlineLevel = 2;
            ws.Cells["A4:A11"].EntireRow.OutlineLevel = 1;
            ws.Cells["A12"].EntireRow.OutlineLevel = 1;
            ws.Cells["A13"].EntireRow.OutlineLevel = 0;
            ws.Cells["A14"].EntireRow.OutlineLevel = 3;
            ws.Cells["A15"].EntireRow.OutlineLevel = 3;
            ws.Cells["A16"].EntireRow.OutlineLevel = 2;
            ws.Cells["A17"].EntireRow.OutlineLevel = 1;
            ws.Cells["A18"].EntireRow.OutlineLevel = 1;
        }
    }
}
