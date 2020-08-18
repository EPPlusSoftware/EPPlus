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
    public class PersonsIntegrationTests : TestBase
    {
        [TestMethod]
        public void PersonCollOnWorkbook()
        {
            using(var package = OpenTemplatePackage("comments.xlsx"))
            {
                var persons = package.Workbook.ThreadedCommentPersons;
                var p = persons.Add("Jan Källman", "Jan Källman", IdentityProvider.NoProvider);
                SaveWorkbook("commentsResult.xlsx", package);
            }
        }

        [TestMethod]
        public void AddPersonToWorkbook()
        {
            using (var package = OpenPackage("commentsWithNewPerson.xlsx", true))
            {
                package.Workbook.Worksheets.Add("test");
                var persons = package.Workbook.ThreadedCommentPersons;
                var p = persons.Add("Jan Källman", "Jan Källman", IdentityProvider.NoProvider);
                package.Save();
            }
        }
    }
}
