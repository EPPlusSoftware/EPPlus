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
                var p = persons.CreateAndAddNewPerson("Jan Källman", "Jan Källman", IdentityProvider.NoProvider);
                SaveWorkbook("commentsResult.xlsx", package);
            }
        }
    }
}
