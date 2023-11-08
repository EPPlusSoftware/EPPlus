using OfficeOpenXml.Export.HtmlExport.Writers.Css;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Xml.Linq;

namespace OfficeOpenXml.Export.HtmlExport.Writers
{
    internal partial class CssTrueWriter : TrueWriterBase
    {
        internal CssTrueWriter(StreamWriter writer) : base(writer)
        {

        }

        internal CssTrueWriter(Stream stream, Encoding encoding): base(stream, encoding)
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

        internal void WriteAndClearCollection(CssRuleCollection collection, bool minify)
        {
            for (int i = 0; i < collection.CssRules.Count(); i++)
            {
                WriteRule(collection[i], minify);
            }

            collection.CssRules.Clear();
        }
    }
}
