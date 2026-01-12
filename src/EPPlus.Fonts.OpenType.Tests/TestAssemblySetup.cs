using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Fonts.OpenType.Tests
{
    [TestClass]
    public class TestAssemblySetup
    {
        [AssemblyInitialize]
        public static void AssemblyInit(TestContext context)
        {
            OpenTypeFonts.ClearFontCache();
        }
    }
}
