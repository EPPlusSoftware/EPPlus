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
            commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y-cell.Size.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
            commands.Add("f");
            AddBorder(cell.BorderData.Top,          cell.LocalPosition.X              , cell.LocalPosition.Y              , cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y              , 0,  2, -2);
            AddBorder(cell.BorderData.Bottom,       cell.LocalPosition.X              , cell.LocalPosition.Y - cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y - cell.Size.Y, 0,  2,  2);
            AddBorder(cell.BorderData.Left,         cell.LocalPosition.X              , cell.LocalPosition.Y              , cell.LocalPosition.X              , cell.LocalPosition.Y - cell.Size.Y, 1,  2, -2);
            AddBorder(cell.BorderData.Right,        cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y              , cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y - cell.Size.Y, 1, -2, -2);
            AddBorder(cell.BorderData.DiagonalDown, cell.LocalPosition.X              , cell.LocalPosition.Y              , cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y - cell.Size.Y, 2,   );
            //AddBorder(cell.BorderData.DiagonalUp,   cell.LocalPosition.X              , cell.LocalPosition.Y - cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y              , 2, );
        }

        private void AddBorder(PdfCellBorderData borderData, double x1, double y1, double x2, double y2, int lt, double doubleOffsetX=0, double doubleOffsetY = 0)
        {
            List<string> commands = new List<string>();
            switch (borderData.BorderStyle)
            {
                case Style.ExcelBorderStyle.None:
                    return;
                case Style.ExcelBorderStyle.Dotted:
                    commands.Add("1.0 w");
                    commands.Add("1 J");
                    commands.Add("[0 2] 0 d");
                    break;
                case Style.ExcelBorderStyle.DashDot:
                    commands.Add("1.0 w");
                    commands.Add("[4 2 1 2] 0 d");
                    break;
                case Style.ExcelBorderStyle.Thin:
                    commands.Add("0.8 w");
                    commands.Add("[] 0 d");
                    break;
                case Style.ExcelBorderStyle.DashDotDot:
                    commands.Add("1.0 w");
                    commands.Add("[4 2 1 2 1 2] 0 d");
                    break;
                case Style.ExcelBorderStyle.Dashed:
                    commands.Add("1.0 w");
                    commands.Add("[4 3] 0 d");
                    break;
                case Style.ExcelBorderStyle.MediumDashDotDot:
                    commands.Add("1.5 w");
                    commands.Add("[6 3 2 3 2 3] 0 d");
                    break;
                case Style.ExcelBorderStyle.MediumDashed:
                    commands.Add("1.5 w");
                    commands.Add("[6 4] 0 d");
                    break;
                case Style.ExcelBorderStyle.MediumDashDot:
                    commands.Add("1.5 w");
                    commands.Add("[6 3 2 3] 0 d");
                    break;
                case Style.ExcelBorderStyle.Thick:
                    commands.Add("2.0 w");
                    commands.Add("[] 0 d");
                    break;
                case Style.ExcelBorderStyle.Medium:
                    commands.Add("1.5 w");
                    commands.Add("[] 0 d");
                    break;
                case Style.ExcelBorderStyle.Double:
                    AddCommand(borderData.BorderColor.ToStrokeCommand());
                    AddCommand("1.0 w");
                    AddCommand("[] 0 d");
                    AddCommand($"{x1.ToPdfString()} {y1.ToPdfString()} m");
                    AddCommand($"{x2.ToPdfString()} {y2.ToPdfString()} l");
                    AddCommand("S");
                    AddCommand("1.0 w");
                    AddCommand("[] 0 d");
                    if (lt==1)
                    {
                        AddCommand($"{(x1 + doubleOffsetX).ToPdfString()} {(y1 + doubleOffsetY).ToPdfString()} m");
                        AddCommand($"{(x2 + doubleOffsetX).ToPdfString()} {(y2 + -doubleOffsetY).ToPdfString()} l");
                    }
                    else
                    {
                        AddCommand($"{(x1 + doubleOffsetX).ToPdfString()} {(y1 + doubleOffsetY).ToPdfString()} m");
                        AddCommand($"{(x2 + -doubleOffsetX).ToPdfString()} {(y2 + doubleOffsetY).ToPdfString()} l");
                    }
                    AddCommand("S");
                    return;
            }
            AddCommand(borderData.BorderColor.ToStrokeCommand());
            foreach (string command in commands)
            {
                AddCommand(command);
            }
            AddCommand($"{x1.ToPdfString()} {y1.ToPdfString()} m");
            AddCommand($"{x2.ToPdfString()} {y2.ToPdfString()} l");
            AddCommand("S");
        }

        public void AddCellContentLayout(PdfCellContentLayout cell, string fontLabel)
        {
            commands.Add("BT");
            commands.Add($"/{fontLabel} {cell.FontData.FontSize.ToPdfString()} Tf");
            commands.Add(cell.FontData.FontColor.ToFillCommand());
            double rot = cell.CellAlignmentData.TextRotation * System.Math.PI / 180.0;
            commands.Add($" {System.Math.Cos(rot).ToPdfString()} {System.Math.Sin(rot).ToPdfString()} {(-System.Math.Sin(rot)).ToPdfString()} {System.Math.Cos(rot).ToPdfString()} {cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} Tm");
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
