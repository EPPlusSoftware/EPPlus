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
        internal Dictionary<ulong, int> _styleCache=new Dictionary<ulong, int>();


        public int Indent { get; set; }


        public void AddAttribute(string attributeName, string attributeValue)
        {
            Require.Argument(attributeName).IsNotNullOrEmpty("attributeName");
            Require.Argument(attributeValue).IsNotNullOrEmpty("attributeValue");
            _attributes.Add(new EpplusHtmlAttribute { AttributeName = attributeName, Value = attributeValue });
        }

        internal void ApplyFormat(bool minify)
        {
            if (minify == false)
            {
                WriteLine();
            }
        }

        internal void ApplyFormatIncreaseIndent(bool minify)
        {
            if (minify==false)
            {
                WriteLine();
                Indent++;
            }
        }

        internal void ApplyFormatDecreaseIndent(bool minify)
        {
            if (minify == false)
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
        internal void SetClassAttributeFromStyle(int styleId, ExcelStyles styles)
        {
            if(styleId <= 0 || styleId >= styles.CellXfs.Count)
            {
                return;
            }
            var xfs = styles.CellXfs[styleId];
            if (xfs.FontId <= 0 && xfs.BorderId <= 0 && xfs.FillId <= 0)
            {
                return;
            }
            var key = (ulong)(xfs.FontId << 32 | xfs.BorderId << 16 | xfs.FillId);
            int id;
            if (_styleCache.ContainsKey(key))
            {
                id = _styleCache[key];
            }
            else
            {
                id = _styleCache.Count + 1;
                _styleCache.Add(key, id);
            }

            AddAttribute("class", $"s{id}");
        }

    }
}
