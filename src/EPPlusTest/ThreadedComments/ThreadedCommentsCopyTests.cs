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
    public class ThreadedCommentsCopyTests : TestBase
    {
        static ExcelPackage _pck;
        [ClassInitialize]
        public static void Init(TestContext context)
        {
            InitBase();
            _pck = OpenPackage("ThreadedCommentCopy.xlsx", true);
        }
        [ClassCleanup]
        public static void Cleanup()
        {
            SaveAndCleanup(_pck);
        }
        [TestMethod]
        public void ShouldCopyThreadedCommentWithinSheet()
        {
            var sheet = _pck.Workbook.Worksheets.Add("WithinSheet");
            var person = sheet.ThreadedComments.Persons.Add("John Doe");
            var person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
            var thread = sheet.Cells["A1"].AddThreadedComment();
            var c1 = thread.AddComment(person2.Id, "Hello");
            var c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

            sheet.Cells[1, 1].Copy(sheet.Cells["A3"]);
            thread = sheet.Cells[3, 1].ThreadedComment;

            Assert.AreEqual(2, thread.Comments.Count);
            Assert.AreEqual("A3", thread.Comments[0].Ref);
            Assert.AreEqual("A3", thread.Comments[1].Ref);
            Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
            Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
        }
        [TestMethod]
        public void ShouldCopyThreadedCommentToNewSheet()
        {
            var sheet = _pck.Workbook.Worksheets.Add("NewSheet_Source");
            var person = sheet.ThreadedComments.Persons.Add("John Doe");
            var person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
            var thread = sheet.Cells["A1"].AddThreadedComment();
            var c1 = thread.AddComment(person2.Id, "Hello");
            var c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

            var sheet2 = _pck.Workbook.Worksheets.Add("NewSheet_Dest");
            sheet.Cells[1, 1].Copy(sheet2.Cells["A3"]);
            thread = sheet2.Cells[3, 1].ThreadedComment;

            Assert.AreEqual(2, thread.Comments.Count);
            Assert.AreEqual("A3", thread.Comments[0].Ref);
            Assert.AreEqual("A3", thread.Comments[1].Ref);
            Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
            Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
        }
        [TestMethod]
        public void ShouldCopyThreadedCommentToNewPackage()
        {
            using (var p = new ExcelPackage())
            {
                var sheet = p.Workbook.Worksheets.Add("test");
                var person = sheet.ThreadedComments.Persons.Add("John Doe");
                var person2 = sheet.ThreadedComments.Persons.Add("Jane Doe");
                var thread = sheet.Cells["A1"].AddThreadedComment();
                var c1 = thread.AddComment(person2.Id, "Hello");
                var c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);
                using (var pck2 = new ExcelPackage())
                {
                    var sheet2 = pck2.Workbook.Worksheets.Add("test2");
                    sheet.Cells[1, 1].Copy(sheet2.Cells["A3"]);
                    thread = sheet2.Cells[3, 1].ThreadedComment;

                    Assert.AreEqual(2, thread.Comments.Count);
                    Assert.AreEqual("A3", thread.Comments[0].Ref);
                    Assert.AreEqual("A3", thread.Comments[1].Ref);
                    Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
                    Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
                    SaveWorkbook("ThreadedCommentCopy_NewPackage.xlsx", pck2);
                }
            }
        }
        [TestMethod]
        public void ShouldCopyWorksheetWithThreadedComment()
        {
            var sheetToCopy = _pck.Workbook.Worksheets.Add("WorksheetCopy_Source");
            var person = sheetToCopy.ThreadedComments.Persons.Add("John Doe");
            var person2 = sheetToCopy.ThreadedComments.Persons.Add("Jane Doe");
            var thread = sheetToCopy.Cells["A1"].AddThreadedComment();
            var c1 = thread.AddComment(person2.Id, "Hello");
            var c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

            var copy = _pck.Workbook.Worksheets.Add("WorksheetCopy_Dest", sheetToCopy);
            thread = copy.Cells[1, 1].ThreadedComment;

            Assert.AreEqual(2, thread.Comments.Count);
            Assert.AreEqual("A1", thread.Comments[0].Ref);
            Assert.AreEqual("A1", thread.Comments[1].Ref);
            Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
            Assert.AreEqual(1, thread.Comments[1].Mentions.Count());
        }
        [TestMethod]
        public void ShouldCopyWorksheetWithThreadedCommentToNewPackage()
        {
            using (var p = new ExcelPackage())
            {
                var sheetToCopy = p.Workbook.Worksheets.Add("WorksheetCopy_Source");
                var person = sheetToCopy.ThreadedComments.Persons.Add("John Doe");
                var person2 = sheetToCopy.ThreadedComments.Persons.Add("Jane Doe");
                var thread = sheetToCopy.Cells["A1"].AddThreadedComment();
                var c1 = thread.AddComment(person2.Id, "Hello");
                var c2 = thread.AddComment(person.Id, "Hello {0}, how are you?", person2);

                using (var pck2 = new ExcelPackage())
                {
                    var copy = pck2.Workbook.Worksheets.Add("WorksheetCopy_Desc", sheetToCopy);
                    thread = copy.Cells[1, 1].ThreadedComment;

                    Assert.AreEqual(2, thread.Comments.Count);
                    Assert.AreEqual("A1", thread.Comments[0].Ref);
                    Assert.AreEqual("A1", thread.Comments[1].Ref);
                    Assert.AreEqual("Hello @Jane Doe, how are you?", c2.Text);
                    Assert.AreEqual(1, thread.Comments[1].Mentions.Count());

                    SaveWorkbook("ThreadedCommentWorksheetCopy_NewPackage.xlsx", pck2);
                }
            }
        }
    }
}
