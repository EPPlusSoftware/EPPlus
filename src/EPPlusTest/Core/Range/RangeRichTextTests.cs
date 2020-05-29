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
    }
}
