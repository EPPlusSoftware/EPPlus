using EPPlus.Fonts.OpenType;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusTest
{
    public class FontTestBase : TestBase
    {
        protected string FontFolder
        {
            get
            {
                return Path.Combine(AppContext.BaseDirectory, "Fonts");
            }
        }
    }
}
