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

using OfficeOpenXml.Export.HtmlExport.CssCollections;
using System;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Writers
{
    internal partial class CssWriter : BaseWriter
    {
        internal CssWriter(StreamWriter writer) : base(writer)
        {

        }

        internal CssWriter(Stream stream): base(stream)
        {

        }

        internal CssWriter(Stream stream, Encoding encoding): base(stream, encoding)
        {

        }

        internal void WriteCssItem(string value, bool minify)
        {
            if (minify)
            {
                _writer.Write(value);
            }
            else
            {
                WriteIndent();
                _writer.WriteLine(value);
            }
        }

        internal void WriteSelectorOpening(string selector, bool minify)
        {
            WriteClass($"{selector}{{", minify);
        }

        internal void WritePropertyDeclaration(Declaration declaration, bool minify)
        {
            WriteCssItem($"{declaration.Name}:{declaration.ValuesToString()};", minify);
        }

        internal void WriteRule(CssRule rule, bool minify)
        {
            WriteSelectorOpening(rule.Selector, minify);

            for (int i = 0; i < rule.Declarations.Count; i++)
            {
                WritePropertyDeclaration(rule.Declarations[i], minify);
            }

            WriteClassEnd(minify);
        }

        internal void WriteAndClearFlush(CssRuleCollection collection, bool minify)
        {
            for (int i = 0; i < collection.CssRules.Count(); i++)
            {
                WriteRule(collection[i], minify);
            }

            collection.CssRules.Clear();
            _writer.Flush();
        }
    }
}
