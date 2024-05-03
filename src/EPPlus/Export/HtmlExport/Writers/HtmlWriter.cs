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
using OfficeOpenXml.Export.HtmlExport.HtmlCollections;
using OfficeOpenXml.Export.HtmlExport.Writers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml.Utils;
using System.Collections.Generic;
using System.IO;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class HtmlWriter : BaseWriter
    {
        internal HtmlWriter(Stream stream, Encoding encoding) : base(stream, encoding)
        {
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

        public void RenderBeginTag(string elementName, List<EpplusHtmlAttribute> attributes = null, bool closeElement = false)
        {
            _newLine = false;
            // avoid writing indent characters for a hyperlinks or images inside a td element
            if (elementName != HtmlElements.A && elementName != HtmlElements.Img)
            {
                WriteIndent();
            }
            _writer.Write($"<{elementName}");


            if (attributes != null)
            {
                foreach (var attribute in attributes)
                {
                    _writer.Write($" {attribute.AttributeName}=\"{attribute.Value}\"");
                }
                attributes.Clear();
            }

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
            RenderBeginTag(element.ElementName, element._attributes, element.IsVoidElement);

            if (element.IsVoidElement)
            {
                if(element.ElementName != HtmlElements.Img)
                {
                    ApplyFormat(minify);
                }
                return;
            }

            if (element._childElements.Count > 0)
            {
                var name = element.ElementName;
                bool noIndent = minify == true ? true : HtmlElements.NoIndentElements.Contains(name);

                ApplyFormatIncreaseIndent(noIndent);

                foreach (var child in element._childElements)
                {
                    RenderHTMLElement(child, minify);
                }

                if(noIndent == false)
                {
                    Indent--;
                }
            }

            Write(element.Content);

            RenderEndTag(element.ElementName);
            if(element.ElementName != "a")
            {
                ApplyFormat(minify);
            }
        }
    }
}
