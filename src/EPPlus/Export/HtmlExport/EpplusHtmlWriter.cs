/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.Style;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class EpplusHtmlWriter : TrueWriterBase
    {
        internal EpplusHtmlWriter(Stream stream, Encoding encoding) : base(stream, encoding)
        {
        }

        private readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();

        public void AddAttribute(string attributeName, string attributeValue)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            Require.Argument(attributeValue).IsNotNullOrEmpty("attributeValue");
            _attributes.Add(new EpplusHtmlAttribute { AttributeName = attributeName, Value = attributeValue });
        }
        public void RenderBeginTag(string elementName, bool closeElement = false)
        {
            _newLine = false;
            // avoid writing indent characters for a hyperlinks or images inside a td element
            if(elementName != HtmlElements.A && elementName != HtmlElements.Img)
            {
                WriteIndent();
            }
            _writer.Write($"<{elementName}");
            foreach (var attribute in _attributes)
            {
                _writer.Write($" {attribute.AttributeName}=\"{attribute.Value}\"");
            }
            _attributes.Clear();

            if (closeElement)
            {
                _writer.Write("/>");
                _writer.Flush();
            }
            else
            {
                _writer.Write(">");
                _elementStack.Push(elementName);
            }
        }

        public void RenderEndTag()
        {
            if (_newLine)
            {
                WriteIndent();
            }

            var elementName = _elementStack.Pop();
            _writer.Write($"</{elementName}>");
            _writer.Flush();
        }

        public void RenderEndTag(string elementName)
        {
            if (_newLine)
            {
                WriteIndent();
            }

            _writer.Write($"</{elementName}>");
            _writer.Flush();
        }

        public void RenderBeginTag(string elementName, List<EpplusHtmlAttribute> attributes, bool closeElement = false)
        {
            _newLine = false;
            // avoid writing indent characters for a hyperlinks or images inside a td element
            if (elementName != HtmlElements.A && elementName != HtmlElements.Img)
            {
                WriteIndent();
            }
            _writer.Write($"<{elementName}");
            foreach (var attribute in attributes)
            {
                _writer.Write($" {attribute.AttributeName}=\"{attribute.Value}\"");
            }
            attributes.Clear();

            if (closeElement)
            {
                _writer.Write("/>");
                _writer.Flush();
            }
            else
            {
                _writer.Write(">");
            }
        }

        public void RenderHTMLElement(HTMLElement element, bool minify)
        {
            if(element._childElements.Count > 0)
            {
                RenderBeginTag(element.ElementName, element._attributes);

                var name = element.ElementName;
                bool noIndent = minify == true ?
                    true :
                    name == HtmlElements.TableData ||
                    name == HtmlElements.TFoot ||
                    name == HtmlElements.TableHeader ||
                    name == HtmlElements.A ||
                    name == HtmlElements.Img;

                ApplyFormatIncreaseIndent(noIndent);

                foreach(var child in element._childElements)
                {
                    RenderHTMLElement(child, minify);
                }

                if(noIndent == false)
                {
                    Indent--;
                }

                RenderEndTag(element.ElementName);
            }
            else
            {
                RenderBeginTag(element.ElementName, element._attributes, true);
            }
            ApplyFormat(minify);
        }

        public void RenderOrderedDict(OrderedDictionary elements, bool minify)
        {
            //if (element._childElements.Count > 0)
            //{
            //    RenderBeginTag(element.ElementName, element._attributes);

            //    var name = element.ElementName;
            //    bool noIndent = minify == true ?
            //        true :
            //        name == HtmlElements.TableData ||
            //        name == HtmlElements.TFoot ||
            //        name == HtmlElements.TableHeader ||
            //        name == HtmlElements.A ||
            //        name == HtmlElements.Img;

            //    ApplyFormatIncreaseIndent(noIndent);

            //    foreach (var child in element._childElements)
            //    {
            //        RenderHTMLElement(child, minify);
            //    }

            //    if (noIndent == false)
            //    {
            //        Indent--;
            //    }

            //    RenderEndTag(element.ElementName);
            //}
            //else
            //{
            //    RenderBeginTag(element.ElementName, element._attributes, true);
            //}
            //ApplyFormat(minify);
        }
    }
}
