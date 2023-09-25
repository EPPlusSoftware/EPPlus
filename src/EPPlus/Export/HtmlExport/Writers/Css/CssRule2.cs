using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Writers.Css
{
    internal class CssRule2
    {
        public CssRule2(string selector)
        {
            Selector = selector;
        }
        public string Selector { get; set; }

        public List<Declaration> Declarations { get; set; }
    }
}
