using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OfficeOpenXml.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest.GlobalConfiguration
{
#if (Core)
    [TestClass]
    public class ConfigurePackageTests
    {
        [TestMethod]
        public void ShouldSetSuppressInitExceptions()
        {
            lock(typeof(ExcelPackage))
            {

                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    ExcelPackage.Configure(x => { 
                        x.SuppressInitializationExceptions = true;
                        x.JsonConfigFileName = "asdf";
                        x.JsonConfigBasePath = "JKLÖ";
                    });
                    using(var package = new ExcelPackage())
                    {
                        Assert.IsTrue(package.InitializationErrors.Count() > 0);
                        Assert.AreEqual(1, package.InitializationErrors.Count());
                    }
                }
                finally
                {
                    ExcelPackage.Configure(x => x.Reset());
                }
            }
        }

        [TestMethod, ExpectedException(typeof(ArgumentException))]
        public void ShouldThrowArgumentExceptionIfErrorsAreNotSuppressed()
        {
            lock(typeof(ExcelPackage))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelPackage.Configure(x => { 
                    x.SuppressInitializationExceptions = false;
                    x.JsonConfigFileName = "asdf";
                    x.JsonConfigBasePath = "JKLÖ";
                });
                try
                {
                    using(var package = new ExcelPackage())
                    {
                        Assert.IsTrue(package.InitializationErrors.Count() > 0);
                        Assert.AreEqual(1, package.InitializationErrors.Count());
                    }
                }
                finally
                {
                    ExcelPackage.Configure(x => x.Reset());
                }
            
            }
        }
    }
#endif
}
