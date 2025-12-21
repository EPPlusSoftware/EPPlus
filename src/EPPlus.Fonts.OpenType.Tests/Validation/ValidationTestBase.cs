using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Validation
{
    public abstract class ValidationTestBase
    {
        public string FontFolder => ValidationTestHelper.FontFolder;
        public List<string> FontFolders => ValidationTestHelper.FontFolders;
    }
}