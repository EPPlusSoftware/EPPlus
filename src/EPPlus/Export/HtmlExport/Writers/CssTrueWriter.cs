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
    internal class CssTrueWriter : TrueWriterBase
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

        ///// <summary>
        ///// write e.g. ".{name}{" to file
        ///// </summary>
        ///// <param name="name"></param>
        ///// <param name="minify"></param>
        //internal void WriteCssClassOpening(string name, bool minify)
        //{
        //    WriteClass($".{name}{{", minify);
        //}

        //internal void WriteSelectorOpening(Selector selector, bool minify) 
        //{
        //    WriteClass($"{selector.Name}{{", minify);
        //}

        internal void WriteSelectorOpening(string selector, bool minify)
        {
            WriteClass($"{selector}{{", minify);
        }

        internal void WritePropertyDeclaration(Declaration declaration, bool minify)
        {
            WriteCssItem($"{declaration.Name}: {declaration.ValuesToString()};", minify);
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

        //internal void WriteRule(Selector selector, bool minify)
        //{
        //    WriteSelectorOpening(selector, minify);

        //    for(int i = 0; i <= selector.Declarations.Count; i++)
        //    {
        //        WritePropertyDeclaration(selector.Declarations[i], minify);
        //    }

        //    WriteClassEnd(minify);
        //}

        //internal void WriteRule(CssRule2 rule2, bool minify)
        //{
        //    WriteSelectorOpening(rule2.Selector, minify);

        //    for (int i = 0; i <= rule2.Declarations.Count; i++)
        //    {
        //        WritePropertyDeclaration(rule2.Declarations[i], minify);
        //    }

        //    WriteClassEnd(minify);
        //}


        ///// <summary>
        ///// Write for example .intro.middle if classes = {"intro", "middle"}
        ///// </summary>
        ///// <param name="classes"></param>
        ///// <param name="minify"></param>
        //internal void WriteSpecifiedCssClassOpening(List<string> classes, bool minify)
        //{
        //    string result = "";
        //    for(int i = 0; i <= classes.Count(); i++)
        //    {
        //        result += $".{classes[i]}";
        //    }

        //    WriteCssClassOpening(result, minify);
        //}




        //internal void WriteWholeCssClass(List<string> classes, List<Tuple<string, List<string>>> cssProperties, bool minify)
        //{
        //    WriteSpecifiedCssClassOpening(classes, minify);
        //    for (int i = 0; i < cssProperties.Count(); i++)
        //    {
        //        WriteCssPropertyMultipleValues(cssProperties[i].Item1, cssProperties[i].Item2, minify);
        //    }

        //    WriteClassEnd(minify);
        //}
    }
}
