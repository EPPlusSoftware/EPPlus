using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.Core
{
    [TestClass]
    public class RichDataTests : TestBase
    {
        [ClassInitialize]
        public static void Init(TestContext context)
        {
        }

        [TestMethod]
        public void RichDataReadTest()
        {
            using (var p = OpenTemplatePackage("RichData.xlsx"))
            {
               Assert.AreEqual(10, p.Workbook.RichData.ValueTypes.Global.Count);
               Assert.AreEqual(3, p.Workbook.RichData.Structures.StructureItems.Count);
               Assert.AreEqual(4, p.Workbook.RichData.Values.Items.Count);
            }
        }
    }
}
