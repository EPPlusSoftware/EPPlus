/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  11/07/2021         EPPlus Software AB       Added Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
    internal partial class EpplusHtmlWriter
    {
#if !NET35 && !NET40
        internal async Task ApplyFormatAsync(bool formatHtml)
        {
            if (formatHtml)
            {
                await WriteLineAsync();
            }
        }

        internal async Task ApplyFormatIncreaseIndentAsync(bool formatHtml)
        {
            if (formatHtml)
            {
                await WriteLineAsync();
                Indent++;
            }
        }

        internal async Task ApplyFormatDecreaseIndentAsync(bool formatHtml)
        {
            if (formatHtml)
            {
                await WriteLineAsync();
                Indent--;
            }
        }

        private async Task WriteIndentAsync()
        {
            for (var x = 0; x < Indent; x++)
            {
                await _writer.WriteAsync(IndentWhiteSpace);
            }
        }

        public async Task RenderBeginTagAsync(string elementName, bool closeElement = false)
        {
            _newLine = false;
            await WriteIndentAsync();
            await _writer.WriteAsync($"<{elementName}");
            foreach (var attribute in _attributes)
            {
                await _writer.WriteAsync($" {attribute.AttributeName}=\"{attribute.Value}\"");
            }
            _attributes.Clear();

            if (closeElement)
            {
                await _writer.WriteAsync("/>");
                await _writer.FlushAsync();
            }
            else
            {
                await _writer.WriteAsync(">");
            }
            _elementStack.Push(elementName);
        }

        public async Task RenderEndTagAsync()
        {
            if (_newLine)
            {
                await WriteIndentAsync();
            }

            var elementName = _elementStack.Pop();
            await _writer.WriteAsync($"</{elementName}>");
            await _writer.FlushAsync();
        }

        public async Task WriteLineAsync()
        {
            _newLine = true;
            await _writer.WriteLineAsync();
        }

        public async Task WriteAsync(string text)
        {
            await _writer.WriteAsync(text);
        }
#endif
    }
}
