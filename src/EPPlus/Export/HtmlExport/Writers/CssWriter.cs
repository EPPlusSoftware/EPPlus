using OfficeOpenXml.Export.HtmlExport.CssCollections;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime;
using System.Text;
using System.Xml.Linq;

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
