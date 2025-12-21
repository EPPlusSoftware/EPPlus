using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Fonts.OpenType.Tests.Serialization
{
    public abstract class SerializationTestBase
    {
        public string FontFolder => SerializationTestHelper.FontFolder;

        public List<string> FontFolders => SerializationTestHelper.FontFolders;
    }
}
