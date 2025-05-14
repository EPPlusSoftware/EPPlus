using OfficeOpenXml.PDF.PdfGraphics;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfContentStream : PdfObject
    {
        private readonly List<string> commands = new List<string>();

        public PdfContentStream(int objectNumber, int version = 0) 
            : base(objectNumber, version)
        {
        }

        public void AddCommand(string command)
        {
            commands.Add(command);
        }

        public void AddText(string fontResourceName, int fontSize, float x, float y, string text)
        {
            commands.Add("BT");
            commands.Add($"/{fontResourceName}{fontSize} Tf");
            commands.Add($"{x} {y} Td");
            commands.Add($"({FixEscapeCharacters(text)}) Tj");
            commands.Add("ET");
        }

        public void AddRectangle(float x, float y, float width, float height, bool stroke = false, bool fill = false, PdfColor strokeColor = null, PdfColor fillColor = null)
        {
            if (stroke != false && strokeColor != null)
            {
                strokeColor.ToStrokeCommand();
            }
            if (fill != false && fillColor != null)
            {
                fillColor.ToFillCommand();
            }
            commands.Add($"{x} {y} {width} {height} re");
            if (fill && stroke)
            {
                commands.Add("B");
            }
            else if (fill)
            {
                commands.Add("f");
            }
            else
            {
                commands.Add("S");
            }
        }

        internal override string RenderDictionary()
        {
            var content = string.Join("\n", commands.ToArray()) + "\n";
            var bytes = Encoding.ASCII.GetBytes(content);
            return $"{objectNumber} {version} obj\n" +
                   $"<< /Length {bytes.Length} >>\n" +
                   $"stream\n{content}endstream\nendobj\n";
        }

        private string FixEscapeCharacters(string text)
        {
            return text.Replace(@"\", @"\\").Replace(@"(", @"\(").Replace(@")", @"\)");
        }
    }
}
