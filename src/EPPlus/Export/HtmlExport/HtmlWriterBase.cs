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
#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport
{
#if !NET35 && !NET40
    internal abstract partial class HtmlWriterBase
    {
        public async Task WriteLineAsync()
        {
            _newLine = true;
            await _writer.WriteLineAsync();
        }

        public async Task WriteAsync(string text)
        {
            await _writer.WriteAsync(text);
        }

        internal protected async Task WriteIndentAsync()
        {
            for (var x = 0; x < Indent; x++)
            {
                await _writer.WriteAsync(IndentWhiteSpace);
            }
        }
        internal async Task ApplyFormatAsync(bool minify)
        {
            if (minify == false)
            {
                await WriteLineAsync();
            }
        }

        internal async Task ApplyFormatIncreaseIndentAsync(bool minify)
        {
            if (minify == false)
            {
                await WriteLineAsync();
                Indent++;
            }
        }

        internal async Task ApplyFormatDecreaseIndentAsync(bool minify)
        {
            if (minify == false)
            {
                await WriteLineAsync();
                Indent--;
            }
        }
        internal async Task WriteClassAsync(string value, bool minify)
        {
            if (minify)
            {
                await _writer.WriteAsync(value);
            }
            else
            {
                await _writer.WriteLineAsync(value);
                Indent = 1;
            }
        }
        internal async Task WriteClassEndAsync(bool minify)
        {
            if (minify)
            {
                await _writer.WriteAsync("}");
            }
            else
            {
                await _writer.WriteLineAsync("}");
                Indent = 0;
            }
        }
        internal async Task WriteCssItemAsync(string value, bool minify)
        {
            if (minify)
            {
                await _writer.WriteAsync(value);
            }
            else
            {
                await WriteIndentAsync();
                _writer.WriteLine(value);
            }
        }
    }
#endif
}