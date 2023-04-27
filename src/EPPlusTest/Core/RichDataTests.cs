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
                
                Assert.AreEqual(2, p.Workbook.Metadata.MetadataTypes.Count);
                Assert.AreEqual(5, p.Workbook.Metadata.FutureMetadataTypes.Count);
                Assert.AreEqual(1, p.Workbook.Metadata.CellMetadata.Count);
                Assert.AreEqual(4, p.Workbook.Metadata.ValueMetadata.Count);

                SaveAndCleanup(p);
            }
        }
    }
}
