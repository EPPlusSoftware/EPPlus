using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeOpenXml.Export.HtmlExport.CssCollections
{
    internal partial class CssRule
    {
        //internal List<SelectorItem> SelectorItems;
        internal List<Declaration> Declarations { get; set; }

        internal string Selector { get; set; }

        internal CssRule(string selector)
        {
            Selector = selector;
            Declarations = new List<Declaration>();
        }

        /// <summary>
        /// Shorthand for ".Declarations.Add(new Declaration(name, values))"
        /// </summary>
        /// <param name="name"></param>
        /// <param name="values"></param>
        internal void AddDeclaration(string name, params string[] values)
        {
            Declarations.Add(new Declaration(name, values));
        }

        internal void AddDeclarationList(List<Declaration> declarations)
        {
            for (int i = 0; i < declarations.Count(); i++)
            {
                Declarations.Add(declarations[i]);
            }
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
