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
    public class RangeRichTextTests
    {
        [TestClass]
        public class DrawingRichTextTests : TestBase
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
        }
    }
}
