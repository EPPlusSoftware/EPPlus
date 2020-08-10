using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ThreadedComments
{
    [TestClass]
    public class ThreadedCommentsInsertTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("ThreadedCommentInsert.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void InsertOneRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("OneRowA1");
            var th=ws.ThreadedComments.Add("A1");
            var p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift down from A1");

            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            ws.InsertRow(1, 1);
            Assert.IsNull(ws.Cells["A1"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["A2"].ThreadedComment);
        }
        [TestMethod]
        public void InsertOneColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("OneColumnA1");
            var th = ws.ThreadedComments.Add("A1");
            var p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift right from A1");

            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            ws.InsertColumn(1, 1);
            Assert.IsNull(ws.Cells["A1"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B1"].ThreadedComment);
        }
        [TestMethod]
        public void InsertTwoRowA1()
        {
            var ws = _pck.Workbook.Worksheets.Add("A1_A2RowA1");
            var th = ws.Cells["A1"].AddThreadedComment();
            var p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift down from A1");

            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            ws.Cells["A1:A2"].Insert(eShiftTypeInsert.Down);
            Assert.IsNull(ws.Cells["A1"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["A3"].ThreadedComment);
        }
        [TestMethod]
        public void InsertTwoColumnA1()
        {
            var ws = _pck.Workbook.Worksheets.Add("A1_B1ColumnA1");
            var th = ws.Cells["A1"].AddThreadedComment();
            var p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift right from A1");

            Assert.IsNotNull(ws.Cells["A1"].ThreadedComment);
            ws.Cells["A1:B1"].Insert(eShiftTypeInsert.Right);
            Assert.IsNull(ws.Cells["A1"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["C1"].ThreadedComment);
        }
        [TestMethod]
        public void InsertInRangeColumn()
        {
            var ws = _pck.Workbook.Worksheets.Add("ColumnInRange");
            var th = ws.Cells["B2:B4"].AddThreadedComment();
            var p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift right from B2");
            ws.ThreadedComments["B3"].AddComment(p.Id, "No shift from B3");
            ws.Cells["B4"].ThreadedComment.AddComment(p.Id, "No shift from B4");

            Assert.IsNotNull(ws.Cells["B2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B3"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B4"].ThreadedComment);
            ws.Cells["A2:B2"].Insert(eShiftTypeInsert.Right);
            Assert.IsNull(ws.Cells["B2"].ThreadedComment);

            Assert.IsNotNull(ws.Cells["D2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B3"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B4"].ThreadedComment);
        }
        [TestMethod]
        public void InsertInRangeRow()
        {
            var ws = _pck.Workbook.Worksheets.Add("RowInRange");
            var th = ws.Cells["B2:D2"].AddThreadedComment();
            var p = ws.ThreadedComments.Persons.Add("Jan Källman");
            th.AddComment(p.Id, "Shift down from B2");
            ws.ThreadedComments["C2"].AddComment(p.Id, "No shift from C2");
            ws.Cells["D2"].ThreadedComment.AddComment(p.Id, "No shift from D2");

            Assert.IsNotNull(ws.Cells["B2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["C2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["D2"].ThreadedComment);
            ws.Cells["B1:B2"].Insert(eShiftTypeInsert.Down);

            Assert.IsNull(ws.Cells["B2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["B4"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["C2"].ThreadedComment);
            Assert.IsNotNull(ws.Cells["D2"].ThreadedComment);
        }
    }
}
