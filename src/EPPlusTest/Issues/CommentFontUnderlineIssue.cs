using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Issues
{
    [TestClass]
    public  class CommentFontUnderlineIssue : TestBase
    {
        [TestMethod]
        public void TestCopyComment()
        {
            using (var pck = OpenTemplatePackage("TestStyledComment.xlsx"))
            {
                // Get the comment object from the worksheet
                var wks = pck.Workbook.Worksheets["Sheet1"];
                var comment1 = wks.Comments[0];

                // Check that the comment in B2 has a custom style
                Assert.AreEqual("B2", comment1.Address);
                Assert.AreEqual("dcf0ff", comment1.BackgroundColor.Name);
                Assert.AreEqual(true, comment1.AutoFit);
                Assert.AreEqual("Tahoma", comment1.Font.FontName);
                Assert.AreEqual(9, comment1.Font.Size);
                Assert.AreEqual(true, comment1.Font.Bold);
                Assert.AreEqual(true, comment1.Font.Italic);
                Assert.AreEqual(true, comment1.Font.UnderLine);
                Assert.AreEqual("0", comment1.Font.Color.Name);

                // Copy the comment from B2 to A2 (also checking that this works when copying a range)
                wks.Cells["B1:B3"].Copy(wks.Cells["A1:A3"]);

                // Check the comment is copied with all properties intact
                var comment2 = wks.Comments[1];
                Assert.AreEqual("A2", comment2.Address);
                Assert.AreEqual(comment1.BackgroundColor.Name, comment2.BackgroundColor.Name);
                Assert.AreEqual(comment1.AutoFit, comment2.AutoFit);
                Assert.AreEqual(comment1.Font.FontName, comment2.Font.FontName);
                Assert.AreEqual(comment1.Font.Size, comment2.Font.Size);
                Assert.AreEqual(comment1.Font.Bold, comment2.Font.Bold);
                Assert.AreEqual(comment1.Font.Italic, comment2.Font.Italic);
                Assert.AreEqual(comment1.Font.UnderLine, comment2.Font.UnderLine);
                Assert.AreEqual(comment1.Font.Color.Name, comment2.Font.Color.Name);
            }
        }
    }
}
