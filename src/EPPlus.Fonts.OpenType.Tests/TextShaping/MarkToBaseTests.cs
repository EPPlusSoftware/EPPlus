using EPPlus.Fonts.OpenType.FontResolver;
using EPPlus.Fonts.OpenType.Tables.Gpos.Data.Lookups.LookupType4;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Interfaces.Fonts;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests.TextShaping
{
    [TestClass]
    public class MarkToBaseTests : FontTestBase
    {
        public override TestContext? TestContext { get; set; }

        [TestInitialize]
        public void TestSetup()
        {
            OpenTypeFonts.ClearFontCache();
            
        }

        [TestMethod]
        public void MarkToBaseTest()
        {
            var font = TestFolderEngine.LoadFont("EB Garamond", FontSubFamily.Regular, ignoreCache: true);

            var shaper = new TextShaper(TestFolderEngine,font);
            string test = "e\u0301"; // e + combining acute

            var shaped = shaper.Shape(test, ShapingOptions.Full);

            Assert.IsTrue(shaped.Glyphs.Any(x => x.XOffset != 0),
                $"Expected XOffset != 0. Got: {string.Join(", ", shaped.Glyphs.Select(g => $"X={g.XOffset}"))}");
        }
       
    }
}
