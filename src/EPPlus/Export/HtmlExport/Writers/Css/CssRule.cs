using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Writers.Css
{
    internal class CssRule
    {
        //internal List<SelectorItem> SelectorItems;
        internal List<Declaration> Declarations { get; set; }

        internal string Selector { get; set; }

        internal CssRule(string selector)
        {
            Selector = selector;
            Declarations = new List<Declaration>();
        }

        //internal string GetSelector()
        //{
        //    string selector = "";
        //    foreach (SelectorItem item in SelectorItems) 
        //    {
        //        selector += item.Name;
        //    }
        //    return selector;
        //}
    }
}
