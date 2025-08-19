using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.ThreadedComments;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.ThreadedComments
{
    [TestClass]
    public class ThreadedCommentsIntegrationTests : TestBase
    {
        [TestMethod]
        public void PersonCollOnWorkbook()
        {
            using (var package = OpenTemplatePackage("comments.xlsx"))
            {
                var persons = package.Workbook.ThreadedCommentPersons;
                Assert.AreEqual(1, persons.Count);
            }
        }

        [TestMethod]
        public void CommentsInWorksheet()
        {
            using (var package = OpenTemplatePackage("comments.xlsx"))
            {
                var comments = package.Workbook.Worksheets.First().ThreadedComments;
                Assert.AreEqual(1, comments.Threads.Count());
                Assert.AreEqual(2, comments.Threads.ElementAt(0).Comments.Count);
            }
        }

        [TestMethod]
        public void CommentsWithMentions()
        {
            using (var package = OpenTemplatePackage("comment_mentions.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.First();
                var comment = sheet.ThreadedComments["A1"].Comments[5];
                
                //sheet.ThreadedComments["A1"].AddComment("A1", sheet.ThreadedComments.Persons.First().Id, "My threaded comment");
                //sheet.Comments.Add(sheet.Cells["A1"], "test", "Mats");
                //sheet.Cells["A1"].ThreadedComments.Comments.RichText;
                //sheet.Cells["A1"].ThreadedComments.Persons;

                Assert.IsNotNull(comment, "Comment was null");
                Assert.IsNotNull(comment.Author, "Author was null");
                Assert.IsNotNull(comment.Mentions, "Mentions was null");
            }
        }

        [TestMethod]
        public void NewCommentsWithMentions()
        {
            using (var package = OpenTemplatePackage("comment_mentions.xlsx"))
            {
                var sheet = package.Workbook.Worksheets.First();

                var author = sheet.ThreadedComments.Persons.First();
                var matsAlm = sheet.ThreadedComments.Persons[1];
                var janKallman = sheet.ThreadedComments.Persons[2];

                sheet.ThreadedComments.Add("A2").AddComment(author.Id, "Some mentions: {0} and {1}. And {0} again.", matsAlm, janKallman);

                SaveWorkbook("NewCommentMentions.xlsx", package);
            }
        }

        [TestMethod]
        public void CreateNewWorkbook()
        {
            using (var package = OpenPackage("NewCommentsWb.xlsx", true))
            {
                var person1 = package.Workbook.ThreadedCommentPersons.Add("Mats Alm");
                var person2 = package.Workbook.ThreadedCommentPersons.Add("Jan Källman");
                var sheet1 = package.Workbook.Worksheets.Add("test 1");
                var sheet2 = package.Workbook.Worksheets.Add("test 2");

                sheet1.ThreadedComments.Add("A1").AddComment(person1.Id, "Hello");
                sheet1.ThreadedComments["A1"].AddComment(person2.Id, "Hello there");
                sheet1.ThreadedComments["A1"].Comments.Last().EditText("Hello there {0}", person1);
                sheet1.ThreadedComments["A1"].ResolveThread();
                sheet1.ThreadedComments["A1"].ReopenThread();
                sheet2.ThreadedComments.Add("B1").AddComment(person1.Id, "Hello again");
                sheet2.ThreadedComments[new ExcelCellAddress("B1")].AddComment(person2.Id, "Hello {0}!", person1);
                sheet2.ThreadedComments["B1"].Remove(sheet2.ThreadedComments["B1"].Comments[0]);
                sheet2.ThreadedComments["B1"].Remove(sheet2.ThreadedComments["B1"].Comments[0]);

                package.Save();
            }
        }

        //[TestMethod]
        //public void LegacyComments()
        //{
        //    using (var package = OpenPackage("LegacyComments.xlsx", true))
        //    {
        //        //var person1 = package.Workbook.ThreadedCommentPersons.Add("Mats Alm");
        //        //var person2 = package.Workbook.ThreadedCommentPersons.Add("Jan Källman");
        //        var sheet1 = package.Workbook.Worksheets.Add("test 1");
        //        var sheet2 = package.Workbook.Worksheets.Add("test 2");

        //        sheet1.Cells["A1"].AddComment("testing", "Mats");
        //        sheet2.Cells["B1"].AddComment("testing testing", "Mats igen");

        //        package.Save();
        //    }
        //}
    }
}
