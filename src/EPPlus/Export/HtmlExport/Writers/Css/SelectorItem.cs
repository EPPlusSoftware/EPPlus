using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Writers.Css
{
    enum SelectorType
    {
        Element,
        Id,
        Class,
        PseudoClass,
        PseudoElement,
        Attribute,
        Universal
    }

    internal class SelectorItem
    {
        SelectorType Type { get; set; }

        string _name = "";

        internal string Name
        { get 
            { 
                switch (Type) 
                {
                    case SelectorType.Element:
                        return _name;
                    case SelectorType.Id:
                        return "#" + _name;
                    case SelectorType.Class:
                        return "." + _name;
                    default:
                        throw new NotSupportedException("Invalid SelectorType");
                }
            }
            set { _name = value; } 
        }

        internal SelectorItem(string name, SelectorType type)
        {
            Type = type;
            Name = name;
        }
    }
}
