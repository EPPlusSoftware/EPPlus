using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    [TestClass]
    public class OfficePropertiesTests
    {
        
        [TestMethod]
        public void ValidateLong()
        {
            using (var pck=new ExcelPackage())
            {
                var ticks = DateTime.Now.Ticks;
                pck.Workbook.Properties.SetCustomPropertyValue("Timestamp", ticks);
                pck.Workbook.Worksheets.Add("Test");

                pck.Save();

                using(var pck2=new ExcelPackage(pck.Stream))
                {
                    Assert.AreEqual((double)ticks, pck.Workbook.Properties.GetCustomPropertyValue("Timestamp"));
                }
            }
        }
        [TestMethod]
        public void ValidateCaseInsensitiveCustomProperties()
        {
            using (var p = new OfficeOpenXml.ExcelPackage())
            {
                p.Workbook.Worksheets.Add("CustomProperties");
                p.Workbook.Properties.SetCustomPropertyValue("Foo", "Bar");
                p.Workbook.Properties.SetCustomPropertyValue("fOO", "bAR");

                Assert.AreEqual("bAR", p.Workbook.Properties.GetCustomPropertyValue("fOo"));
            }
        }
        [TestMethod]
        public void ValidateCaseInsensitiveCustomProperties_Loading()
        {
            var p = new OfficeOpenXml.ExcelPackage();
            p.Workbook.Worksheets.Add("CustomProperties");
            p.Workbook.Properties.SetCustomPropertyValue("fOO", "bAR");
            p.Workbook.Properties.SetCustomPropertyValue("Foo", "Bar");

            p.Save();

            var p2 = new OfficeOpenXml.ExcelPackage(p.Stream);

            Assert.AreEqual("Bar", p2.Workbook.Properties.GetCustomPropertyValue("fOo"));

            p.Dispose();
            p2.Dispose();
        }
    }
}
