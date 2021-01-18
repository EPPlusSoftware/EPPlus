using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core.Range
{
    [TestClass]
    public class RangeRichTextTests : TestBase
    {
        static ExcelPackage _pck;
        static ExcelWorksheet _ws;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            _pck = OpenPackage("RangeRichText.xlsx", true);
            _ws = _pck.Workbook.Worksheets.Add("Richtext");
        }
        [ClassCleanup]
        public static void Cleanup()
        {

            SaveAndCleanup(_pck);
        }

        [TestMethod]
        public void AddThreeParagraphsAndValidate()
        {
            var r = _ws.Cells["A1"];
            var r1=r.RichText.Add("Line1\n");
            r1.PreserveSpace = true;
            var r2 = r.RichText.Add("Line2\n");
            r2.PreserveSpace = true;
            var r3 = r.RichText.Add("Line3");
            r3.PreserveSpace = true;
            r3.Bold = true;
            r3.Italic = true;
            r3.Size = 19.5F;

            Assert.AreEqual("Line1\nLine2\nLine3", r.Text);
            Assert.AreEqual("Line1\nLine2\nLine3", r.RichText.Text);
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddEmptyStringShouldThrowArgumentException()
        {
            _ws.Cells["D1"].RichText.Add("");
        }
        [TestMethod]
        [ExpectedException(typeof(ArgumentException))]
        public void AddNullShouldThrowArgumentException()
        {
            _ws.Cells["D1"].RichText.Add(null);
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void SettingNameToEmptyStringShouldThrowInvalidOperationException()
        {
            _ws.Cells["D1"].RichText.Add(" ");
            _ws.Cells["D1"].RichText[0].Text = "";
        }
        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        public void SettingTextToNullShouldThrowInvalidOperationException()
        {
            _ws.Cells["D1"].RichText.Add(" ");
            _ws.Cells["D1"].RichText[0].Text = null;
        }
        [TestMethod]
        public void SettingRichTextTextToNullShouldClearRichText()
        {
            _ws.Cells["D1"].RichText.Add(" ");
            Assert.IsTrue(_ws.Cells["D1"].IsRichText);
            _ws.Cells["D1"].RichText.Text = null;
            Assert.AreEqual(0, _ws.Cells["D1"].RichText.Count);
            Assert.IsFalse(_ws.Cells["D1"].IsRichText);
        }
        [TestMethod]
        public void SettingRichTextTextToEmptyStringShouldClearRichText()
        {
            _ws.Cells["D1"].RichText.Add(" ");
            Assert.IsTrue(_ws.Cells["D1"].IsRichText);
            _ws.Cells["D1"].RichText.Text = null;
            Assert.AreEqual(0, _ws.Cells["D1"].RichText.Count);
            Assert.IsFalse(_ws.Cells["D1"].IsRichText);
        }
        [TestMethod]
        public void RemoveVerticalAlign()
        {
            var p=_ws.Cells["G1"].RichText.Add("RemoveVerticalAlign");
            p.VerticalAlign = OfficeOpenXml.Style.ExcelVerticalAlignmentFont.Baseline;
            p.VerticalAlign = OfficeOpenXml.Style.ExcelVerticalAlignmentFont.None;
            Assert.AreEqual(p.VerticalAlign, OfficeOpenXml.Style.ExcelVerticalAlignmentFont.None);
        }
        [TestMethod]
        public void ValidateIsRichTextValuesAndTexts()
        {
            using (var p1 = new ExcelPackage())
            {
                var ws = p1.Workbook.Worksheets.Add("RichText");
                var v = "Player's taunt success & you attack them";
                ws.Cells["A1"].Value = v;
                p1.Save();

                using (var p2 = new ExcelPackage(p1.Stream))
                {
                    Assert.AreEqual(v, ws.Cells["A1"].Value);
                    ws.Cells["A1"].IsRichText = true;
                    Assert.AreEqual(v, ws.Cells["A1"].Value);
                    Assert.AreEqual(v, ws.Cells["A1"].RichText.Text);
                    ws.Cells["A1"].IsRichText = false;
                    Assert.AreEqual(v, ws.Cells["A1"].Value);

                    p2.Save();
                }
            }
        }
        [TestMethod]
        public void ValidateRichTextOverwriteByArray()
        {
            var ws = _pck.Workbook.Worksheets.Add("RichTextOverwriteArray");
            for(int row=1;row<10;row++)
            {
                for (int col = 1; col < 10; col++)
                {
                    ws.Cells[row, col].RichText.Text = $"Cell {ExcelCellBase.GetAddress(row,col)}";
                }
            }

            var array = new object[,] { { "Overwrite cell 1-1", "Overwrite cell 1-2" }, { "Overwrite cell 2-1", "Overwrite cell 2-2" } };

            ws.Cells["C3:D4"].Value = array;

            Assert.IsTrue(ws.Cells["C2"].IsRichText);
            Assert.IsTrue(ws.Cells["E3"].IsRichText);
            Assert.IsTrue(ws.Cells["A4"].IsRichText);
            Assert.IsTrue(ws.Cells["C5"].IsRichText);

            Assert.IsFalse(ws.Cells["C3"].IsRichText);
            Assert.IsFalse(ws.Cells["D4"].IsRichText);

            Assert.AreEqual("Cell C2", ws.Cells["C2"].Value);
            Assert.AreEqual("Overwrite cell 1-1", ws.Cells["C3"].Value);
            Assert.AreEqual("Overwrite cell 2-2", ws.Cells["D4"].Value);
            Assert.AreEqual("Cell A4", ws.Cells["A4"].Value);
            Assert.AreEqual("Cell D5", ws.Cells["D5"].Value);
        }

    }
}
