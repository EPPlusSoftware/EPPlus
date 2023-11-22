using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

#if !NET35 && !NET40
using System.Threading.Tasks;
#endif

namespace OfficeOpenXml.Export.HtmlExport.Writers
{
#if !NET35 && !NET40
    internal abstract partial class BaseWriter
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

        //---

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
