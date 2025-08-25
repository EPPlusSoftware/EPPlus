using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfLayout;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfContentStream : PdfObject
    {
        private readonly List<string> commands = new List<string>();

        public PdfContentStream(int objectNumber, string command = null, int version = 0)
            : base(objectNumber, version)
        {
            if(!string.IsNullOrEmpty(command))
            {
                commands.Add(command);
            }
        }

        public void AddCommand(string command)
        {
            commands.Add(command);
        }

        public void AddText(string fontResourceName, double fontSize, double x, double y, string text)
        {
            commands.Add("BT");
            commands.Add($"/{fontResourceName} {fontSize.ToPdfString()} Tf");
            commands.Add($"{x.ToPdfString()} {y.ToPdfString()} Td");
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
            commands.Add($"{x.ToPdfString()} {y.ToPdfString()} {width.ToPdfString()} {height.ToPdfString()} re");
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

        public void AddCellLayout(PdfCellLayout cell)
        {
            //Need to add pattern commands
            commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
            commands.Add($"{cell.Position.X.ToPdfString()} {cell.Position.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
            commands.Add("f");
            //Top Border
            if(cell.BorderData.Top.BorderStyle != Style.ExcelBorderStyle.None)
            {

            }
            //Bottom Border
            if (cell.BorderData.Bottom.BorderStyle != Style.ExcelBorderStyle.None)
            {

            }
            //Left Border
            if (cell.BorderData.Left.BorderStyle != Style.ExcelBorderStyle.None)
            {

            }
            //Right Border
            if (cell.BorderData.Right.BorderStyle != Style.ExcelBorderStyle.None)
            {

            }
        }

        public void AddCellContentLayout(PdfCellContentLayout cell, string fontLabel)
        {
            commands.Add("BT");
            commands.Add($"/{fontLabel} {cell.FontData.FontSize.ToPdfString()} Tf");
            commands.Add(cell.FontData.FontColor.ToFillCommand());
            double rot = cell.CellAlignmentData.TextRotation * System.Math.PI / 180.0;
            commands.Add($" {System.Math.Cos(rot).ToPdfString()} {System.Math.Sin(rot).ToPdfString()} {(-System.Math.Cos(rot)).ToPdfString()} {System.Math.Cos(rot).ToPdfString()} {cell.Position.X.ToPdfString()} {cell.Position.Y.ToPdfString()} Tm");
            commands.Add($"({FixEscapeCharacters(cell.FontData.Text)}) Tj");
            commands.Add("ET");
        }

        internal override string RenderDictionary()
        {
            var content = string.Join("\n", commands.ToArray()) + "\n";
            var bytes = Encoding.ASCII.GetBytes(content);
            return $"<< /Length {bytes.Length} >>\n" +
                   $"stream\n{content}endstream";
        }

        private string FixEscapeCharacters(string text)
        {
            return text.Replace(@"\", @"\\").Replace(@"(", @"\(").Replace(@")", @"\)");
        }
    }
}
