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
    internal partial class CssRule
    {
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
    }
}
