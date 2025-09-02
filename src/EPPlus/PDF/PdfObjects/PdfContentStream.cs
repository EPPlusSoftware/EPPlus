using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfLayout;
using System.Collections.Generic;
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

        public void AddCellLayout(PdfCellLayout cell)
        {
            if (cell.CellFillData.BackgroundColor.A >= 0.99999f)
            {
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add(cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y - cell.Size.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
            }
        }

        public void AddBorderLayout(PdfCellBorderLayout cell)
        {
            AddBorder(cell.BorderData.Top, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y, 0, 2, -2);
            AddBorder(cell.BorderData.Bottom, cell.LocalPosition.X, cell.LocalPosition.Y - cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y - cell.Size.Y, 0, 2, 2);
            AddBorder(cell.BorderData.Left, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X, cell.LocalPosition.Y - cell.Size.Y, 1, 2, -2);
            AddBorder(cell.BorderData.Right, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y - cell.Size.Y, 1, -2, -2);
            AddBorder(cell.BorderData.DiagonalDown, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y - cell.Size.Y, 2, 2, 2);
            AddBorder(cell.BorderData.DiagonalUp, cell.LocalPosition.X, cell.LocalPosition.Y - cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y, 3, 2, 2);
        }

        private void AddBorder(PdfCellBorderData borderData, double x1, double y1, double x2, double y2, int lt, double doubleOffsetX=0, double doubleOffsetY = 0)
        {
            List<string> commands = new List<string>();
            switch (borderData.BorderStyle)
            {
                case Style.ExcelBorderStyle.None:
                    return;
                case Style.ExcelBorderStyle.Hair:
                    commands.Add("0.5 w");
                    if (lt == 2 || lt == 3)
                        commands.Add("0 J");
                    else
                        commands.Add("2 J");
                    commands.Add("[] 0 d");
                    break;
                case Style.ExcelBorderStyle.Dotted:
                    commands.Add("1.1 w");
                    commands.Add("1 J");
                    commands.Add("[0 2] 0 d");
                    break;
                case Style.ExcelBorderStyle.DashDot:
                    commands.Add("1.1 w");
                    commands.Add("[4 2 1 2] 0 d");
                    break;
                case Style.ExcelBorderStyle.Thin:
                    commands.Add("0.85 w");
                    if (lt == 2 || lt == 3)
                        commands.Add("0 J");
                    else
                        commands.Add("2 J");
                    commands.Add("[] 0 d");
                    break;
                case Style.ExcelBorderStyle.DashDotDot:
                    commands.Add("1.1 w");
                    commands.Add("[4 2 1 2 1 2] 0 d");
                    break;
                case Style.ExcelBorderStyle.Dashed:
                    commands.Add("1.1 w");
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
                    if (lt == 2 || lt == 3)
                        commands.Add("0 J");
                    else
                        commands.Add("2 J");
                    commands.Add("[] 0 d");
                    break;
                case Style.ExcelBorderStyle.Medium:
                    commands.Add("1.5 w");
                    if(lt==2||lt==3)
                        commands.Add("0 J");
                    else
                        commands.Add("2 J");
                    commands.Add("[] 0 d");
                    break;
                case Style.ExcelBorderStyle.SlantDashDot:
                    AddCommand(borderData.BorderColor.ToStrokeCommand());
                    AddCommand("Q");
                    AddCommand("q");
                    AddCommand("1.1 w");
                    if (lt == 2 || lt == 3)
                        AddCommand("0 J");
                    else
                        AddCommand("2 J");
                    AddCommand("[4 2 1 2] 0 d");
                    AddCommand($"1 0 0.6 1 0 0 cm");
                    //calculate new x and y
                    var nx1 = x1 + y1 * 0.6d;
                    var tx1 = nx1 - x1;
                    var nx2 = x2 + y2 * 0.6d;
                    var tx2 = nx2 - x2;
                    AddCommand($"{(x1-tx1).ToPdfString()} {y1.ToPdfString()} m");
                    AddCommand($"{(x2-tx2).ToPdfString()} {y2.ToPdfString()} l");
                    AddCommand("S");
                    AddCommand("Q");
                    AddCommand("q");
                    return;
                case Style.ExcelBorderStyle.Double:
                    AddCommand(borderData.BorderColor.ToStrokeCommand());
                    AddCommand("1.0 w");
                    AddCommand("[] 0 d");
                    if (lt == 2)
                    {
                        AddCommand("0 J");
                        AddCommand($"{((x1 + doubleOffsetX) + doubleOffsetX).ToPdfString()} {((y1 + -doubleOffsetY)).ToPdfString()} m");
                        AddCommand($"{((x2 + -doubleOffsetX)).ToPdfString()} {((y2 + doubleOffsetY) + doubleOffsetY).ToPdfString()} l");
                    }
                    else if(lt==3)
                    {
                        AddCommand("0 J");
                        AddCommand($"{((x1 + doubleOffsetX) + doubleOffsetX).ToPdfString()} {((y1 + doubleOffsetY)).ToPdfString()} m");
                        AddCommand($"{((x2 + -doubleOffsetX)).ToPdfString()} {((y2 + -doubleOffsetY) - doubleOffsetY ).ToPdfString()} l");
                    }
                    else
                    {
                        AddCommand("2 J");
                        AddCommand($"{x1.ToPdfString()} {y1.ToPdfString()} m");
                        AddCommand($"{x2.ToPdfString()} {y2.ToPdfString()} l");
                    }
                    AddCommand("S");
                    AddCommand("1.0 w");
                    AddCommand("[] 0 d");
                    if (lt == 1)
                    {
                        AddCommand("2 J");
                        AddCommand($"{(x1 + doubleOffsetX).ToPdfString()} {(y1 + doubleOffsetY).ToPdfString()} m");
                        AddCommand($"{(x2 + doubleOffsetX).ToPdfString()} {(y2 + -doubleOffsetY).ToPdfString()} l");
                    }
                    else if (lt == 2)
                    {
                        AddCommand("0 J");
                        AddCommand($"{((x1 + doubleOffsetX)).ToPdfString()} {((y1 + -doubleOffsetY) - doubleOffsetY).ToPdfString()} m");
                        AddCommand($"{((x2 + -doubleOffsetX) - doubleOffsetX).ToPdfString()} {(y2 + doubleOffsetY).ToPdfString()} l");
                    }
                    else if(lt == 3)
                    {
                        AddCommand("0 J");
                        AddCommand($"{((x1 + doubleOffsetX)).ToPdfString()} {((y1 + doubleOffsetY) + doubleOffsetY).ToPdfString()} m");
                        AddCommand($"{((x2 + -doubleOffsetX) - doubleOffsetX).ToPdfString()} {((y2 + -doubleOffsetY)).ToPdfString()} l");
                    }
                    else
                    {
                        AddCommand("2 J");
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
            AddCommand($"{x1.ToPdfStringF4()} {y1.ToPdfStringF4()} m");
            AddCommand($"{x2.ToPdfStringF4()} {y2.ToPdfStringF4()} l");
            AddCommand("S");
        }

        public void AddCellContentLayout(PdfCellContentLayout cell, string fontLabel)
        {
            commands.Add("BT");
            commands.Add($"/{fontLabel} {cell.FontData.FontSize.ToPdfString()} Tf");
            commands.Add(cell.FontData.FontColor.ToFillCommand());
            double rot = cell.CellAlignmentData.TextRotation * System.Math.PI / 180.0;
            commands.Add($"{System.Math.Cos(rot).ToPdfString()} {System.Math.Sin(rot).ToPdfString()} {(-System.Math.Sin(rot)).ToPdfString()} {System.Math.Cos(rot).ToPdfString()} {cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} Tm");
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
