﻿/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Writers
{
    internal abstract partial class BaseWriter
    {
        protected readonly Stream _stream;
        protected readonly StreamWriter _writer;

        protected const string IndentWhiteSpace = "  ";
        protected bool _newLine;

        internal protected HashSet<string> _images = new HashSet<string>();

        internal BaseWriter(Stream stream, Encoding encoding)
        {
            _stream = stream;
            _writer = new StreamWriter(stream, encoding);
        }

        internal BaseWriter(Stream stream)
        {
            if (!stream.CanWrite)
            {
                throw new IOException("Parameter stream must be a writable System.IO.Stream");
            }
            _stream = stream;
            _writer = new StreamWriter(stream);
        }

        public BaseWriter(StreamWriter writer)
        {
            _stream = writer.BaseStream;
            _writer = writer;
        }
        internal int Indent { get; set; }

        public void WriteLine()
        {
            _newLine = true;
            _writer.WriteLine();
        }

        public void Write(string text)
        {
            _writer.Write(text);
        }

        internal protected void WriteIndent()
        {
            for (var x = 0; x < Indent; x++)
            {
                _writer.Write(IndentWhiteSpace);
            }
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
            if (minify == false)
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

        internal void WriteClass(string value, bool minify)
        {
            if (minify)
            {
                _writer.Write(value);
            }
            else
            {
                _writer.WriteLine(value);
                Indent = 1;
            }
        }

        internal void WriteClassEnd(bool minify)
        {
            if (minify)
            {
                _writer.Write("}");
            }
            else
            {
                _writer.WriteLine("}");
                Indent = 0;
            }
        }
    }
}
