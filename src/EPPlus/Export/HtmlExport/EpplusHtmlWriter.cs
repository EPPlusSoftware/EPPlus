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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
#if !NET35
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class EpplusHtmlWriter
    {
        public const string IndentWhiteSpace = "  ";
        private bool _newLine;

        public EpplusHtmlWriter(Stream stream)
        {
            _stream = stream;
            _writer = new StreamWriter(stream);
        }

        private readonly Stream _stream;
        private readonly StreamWriter _writer;
        private readonly Stack<string> _elementStack = new Stack<string>();
        private readonly List<EpplusHtmlAttribute> _attributes = new List<EpplusHtmlAttribute>();

        public int Indent { get; set; }


        public void AddAttribute(string attributeName, string attributeValue)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            Require.Argument(attributeValue).IsNotNullOrEmpty("attributeValue");
            _attributes.Add(new EpplusHtmlAttribute { AttributeName = attributeName, Value = attributeValue });
        }

        internal void ApplyFormat(bool formatHtml)
        {
            if (formatHtml)
            {
                WriteLine();
            }
        }

        internal void ApplyFormatIncreaseIndent(bool formatHtml)
        {
            if (formatHtml)
            {
                WriteLine();
                Indent++;
            }
        }

        internal void ApplyFormatDecreaseIndent(bool formatHtml)
        {
            if (formatHtml)
            {
                WriteLine();
                Indent--;
            }
        }

        private void WriteIndent()
        {
            for (var x = 0; x < Indent; x++)
            {
                _writer.Write(IndentWhiteSpace);
            }
        }

        public void RenderBeginTag(string elementName, bool closeElement = false)
        {
            _newLine = false;
            WriteIndent();
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
            }
            _elementStack.Push(elementName);
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

        public void WriteLine()
        {
            _newLine = true;
            _writer.WriteLine();
        }

        public void Write(string text)
        {
            _writer.Write(text);
        }
    }
}
