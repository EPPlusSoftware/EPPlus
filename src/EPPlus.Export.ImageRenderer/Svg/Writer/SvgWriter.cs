using EPPlus.Export.ImageRenderer.Svg.NodeAttributes;
using EPPlus.Export.ImageRenderer.Svg.Nodes;
using OfficeOpenXml.Export.HtmlExport.Writers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Svg.Writer
{
    internal class SvgWriter : BaseWriter
    {
        internal SvgWriter(Stream stream, Encoding encoding) : base(stream, encoding)
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

        public void RenderBeginTag(string elementName, List<SvgAttributeBase> attributes = null, bool closeElement = false)
        {
            _newLine = false;

            WriteIndent();
            //// avoid writing indent characters for a hyperlinks or images inside a td element
            //if (elementName != HtmlElements.A && elementName != HtmlElements.Img)
            //{
            //    WriteIndent();
            //}
            _writer.Write($"<{elementName}");


            if (attributes != null)
            {
                foreach (var attribute in attributes)
                {
                    _writer.Write($" {attribute.Name}=\"{attribute.Value}\"");
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

        public void RenderSvgElementWithoutEndNode(SvgElement element, bool minify)
        {
            RenderBeginTag(element.ElementName, element._attributes, element.IsVoidElement);

            if (element.IsVoidElement)
            {
                ApplyFormat(minify);
                //if (element.ElementName != SvgElements.Img)
                //{
                //ApplyFormat(minify);
                // }
                return;
            }

            if (element._childElements.Count > 0)
            {
                var name = element.ElementName;
                bool noIndent = minify == true ? true : SvgElements.NoIndentElements.Contains(name);

                ApplyFormatIncreaseIndent(noIndent);

                foreach (var child in element._childElements)
                {
                    RenderSvgElement(child, minify);
                }

                if (noIndent == false)
                {
                    Indent--;
                }
            }

            Write(element.Content);

            if (element.ElementName != "a")
            {
                ApplyFormat(minify);
            }
        }

        public void RenderSvgElement(SvgElement element, bool minify)
        {
            //RenderSvgElementWithoutEndNode(element, minify);
            RenderBeginTag(element.ElementName, element._attributes, element.IsVoidElement);

            if (element.IsVoidElement)
            {
                ApplyFormat(minify);
                //if (element.ElementName != SvgElements.Img)
                //{
                //ApplyFormat(minify);
                // }
                return;
            }

            if (element._childElements.Count > 0)
            {
                var name = element.ElementName;
                bool noIndent = minify == true ? true : SvgElements.NoIndentElements.Contains(name);

                ApplyFormatIncreaseIndent(noIndent);

                foreach (var child in element._childElements)
                {
                    RenderSvgElement(child, minify);
                }

                if (noIndent == false)
                {
                    Indent--;
                }
            }

            Write(element.Content);

            RenderEndTag(element.ElementName);
            if (element.ElementName != "a")
            {
                ApplyFormat(minify);
            }
        }
    }
}
