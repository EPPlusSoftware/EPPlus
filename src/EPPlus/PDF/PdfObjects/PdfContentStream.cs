using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
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

        public void AddText(string fontResourceName, double fontSize, double x, double y, string text)
        {
            commands.Add("BT");
            commands.Add($"/{fontResourceName} {PdfString.Convert(fontSize)} Tf");
            commands.Add($"{PdfString.Convert(x)} {PdfString.Convert(y)} Td");
            commands.Add($"({FixEscapeCharacters(text)}) Tj");
            commands.Add("ET");
        }

        public void AddRectangle(double x, double y, double width, double height, bool stroke = false, bool fill = false, PdfColor strokeColor = null, PdfColor fillColor = null)
        {
            if (stroke != false && strokeColor != null)
            {
                commands.Add(strokeColor.ToStrokeCommand());
            }
            if (fill != false && fillColor != null)
            {
                commands.Add(fillColor.ToFillCommand());
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
            return $"<< /Length {bytes.Length} >>\n" +
                   $"stream\n{content}endstream\n";
        }

        private string FixEscapeCharacters(string text)
        {
            return text.Replace(@"\", @"\\").Replace(@"(", @"\(").Replace(@")", @"\)");
        }
    }
}
