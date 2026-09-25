using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.FontScanning
{
    [TestClass]
    public class SystemFontsTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestMethod]
        public void AptosNarrowTest1()
        {
            var font = SystemFontsEngine.LoadFont("Aptos Narrow", FontSubFamily.Regular);
            Assert.AreEqual("Aptos Narrow", font.FullName);
        }
    }
}
