using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Writers.Css
{
    internal class Selector
    {
        public string Name { get; set; }

        public List<Declaration> Declarations { get; set; }

        public Selector(string name) 
        {
            Name = name;
            Declarations = new List<Declaration>();
        }
    }
}
