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
    }
}
