using Microsoft.VisualStudio.TestTools.UnitTesting;
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
                var comments = package.Workbook.Worksheets.First().ThreadedComments;
                var mentions = comments.Threads.First().Comments.ElementAt(5).Mentions;
            }
        }
    }
}
