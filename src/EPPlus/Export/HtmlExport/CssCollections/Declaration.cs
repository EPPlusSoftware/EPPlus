﻿using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.CssCollections
{
    /// <summary>
    /// A css Declaration is the combo of a property and its values.
    /// </summary>
    internal class Declaration
    {
        public string Name { get; set; }
        public List<string> Values { get; set; }

        internal Declaration(string name, params string[] values)
        {
            Name = name;
            Values = new List<string>(values);
        }

        internal string ValuesToString()
        {
            string res = "";

            for (int i = 0; i < Values.Count(); i++)
            {
                res += $"{Values[i]} ";
            }

            return res.TrimEnd();
        }

        internal void AddValues(params string[] values)
        {
            Values.AddRange(values);
        }

    }
}