/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;

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
