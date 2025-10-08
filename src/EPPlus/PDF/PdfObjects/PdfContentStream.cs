using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfLayout;
using OfficeOpenXml.PDF.PdfResources;
using OfficeOpenXml.PDF.PdfSettings;
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

        public void AddCellLayout(PdfCellLayout cell, string label)
        {
            if (cell.CellFillData.GradientFillData != null && cell.CellFillData.PattenStyle != Style.ExcelFillStyle.Solid)
            {
                commands.Add("q");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re W n");
                commands.Add($"/{label} sh");
                commands.Add("f");
                commands.Add("Q");
            }
            else if (cell.CellFillData.BackgroundColor != null && cell.CellFillData.PattenStyle == Style.ExcelFillStyle.Solid)
            {
                commands.Add("q");
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add(cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
                commands.Add("Q");
            }
            else if (cell.CellFillData.BackgroundColor != null && cell.CellFillData.PattenStyle != Style.ExcelFillStyle.None)
            {
                commands.Add("q");
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add(cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
                commands.Add("Q");
                commands.Add("q");
                commands.Add("/Pattern cs");
                commands.Add($"/{label} scn");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("f");
                commands.Add("Q");
            }
        }

        public void AddBorderLayout(PdfCellBorderLayout cell)
        {
            var borderRenderer = new PdfBorderRenderer();
            commands.Add("q");
            borderRenderer.RenderBorder(this, cell.BorderData.Top, LineType.Horizontal, cell.LocalPosition.X, cell.LocalPosition.Y + cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y + cell.Size.Y);
            borderRenderer.RenderBorder(this, cell.BorderData.Bottom, LineType.Horizontal, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y);
            borderRenderer.RenderBorder(this, cell.BorderData.Left, LineType.Vertical, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X, cell.LocalPosition.Y + cell.Size.Y);
            borderRenderer.RenderBorder(this, cell.BorderData.Right, LineType.Vertical, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y + cell.Size.Y);
            borderRenderer.RenderBorder(this, cell.BorderData.DiagonalUp, LineType.DiagonalUp, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y + cell.Size.Y);
            borderRenderer.RenderBorder(this, cell.BorderData.DiagonalDown, LineType.DiagonalDown, cell.LocalPosition.X, cell.LocalPosition.Y + cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y);
            commands.Add("Q");
        }

        public void AddCellContentLayout(PdfCellContentLayout cell, PdfFontResource font)
        {
            commands.Add("q");
            if (cell.Clip)
            {
                commands.Add($"{cell.Clipping.X.ToPdfString()} {cell.Clipping.Y.ToPdfString()} {cell.Clipping.Width.ToPdfString()} {cell.Clipping.Height.ToPdfString()} re W n");
            }
            double size = cell.FontData.FontSize;
            double scale = cell.FontData.FontSize / font.fontData.HeadTable.UnitsPerEm;
            double rot = cell.CellAlignmentData.TextRotation * System.Math.PI / 180.0;
            PDF.Math.Matrix3x3 m1 = new PDF.Math.Matrix3x3(System.Math.Cos(rot), System.Math.Sin(rot), -System.Math.Sin(rot), System.Math.Cos(rot), cell.LocalPosition.X, cell.LocalPosition.Y);
            if (cell.FontData.Bold)
            {
                commands.Add("0.25 w");
                commands.Add("2 Tr");
            }
            if (cell.FontData.Italic)
            {
                var i = font.fontData.postTable.italicAngle;
                if (i <= 0) i = 12d * System.Math.PI / 180.0d;
                PDF.Math.Matrix3x3 m2 = PDF.Math.Matrix3x3.Identity;
                m2.C = System.Math.Tan(i);
                m1 = m2 * m1;
            }
            if (cell.FontData.SuperScript)
            {
                var supOffX = font.fontData.Os2Table.ySuperscriptXOffset * scale;
                var supOffY = font.fontData.Os2Table.ySuperscriptYOffset * scale;
                var supSizeX = font.fontData.Os2Table.ySuperscriptXSize * scale;
                var supSizeY = font.fontData.Os2Table.ySuperscriptYSize * scale;
                m1 = new PDF.Math.Matrix3x3(m1.A, m1.B, m1.C, m1.D, m1.E + supOffX, m1.F + supOffY);
                size = supSizeY;
            }
            if (cell.FontData.SubScript)
            {
                var supOffX = font.fontData.Os2Table.ySubscriptXOffset * scale;
                var supOffY = font.fontData.Os2Table.ySubscriptYOffset * scale;
                var supSizeX = font.fontData.Os2Table.ySubscriptXSize * scale;
                var supSizeY = font.fontData.Os2Table.ySubscriptYSize * scale;
                m1 = new PDF.Math.Matrix3x3(m1.A, m1.B, m1.C, m1.D, m1.E + supOffX, m1.F + supOffY);
                size = supSizeY;
            }
            commands.Add("BT");
            commands.Add($"/{font.Label} {size.ToPdfString()} Tf");
            commands.Add(cell.FontData.FontColor.ToFillCommand());
            commands.Add($"{m1.A.ToPdfString()} {m1.B.ToPdfString()} {m1.C.ToPdfString()} {m1.D.ToPdfString()} {m1.E.ToPdfString()} {m1.F.ToPdfString()} Tm");
            commands.Add($"({FixEscapeCharacters(cell.FontData.Text)}) Tj");
            commands.Add("ET");
            if (cell.FontData.Underline)
            {
                var underlinePos = font.fontData.postTable.underlinePosition * scale;
                var underlineWidth = font.fontData.postTable.underlineThickness * scale;
                var start = m1.Transform(new Math.Vector2(0, underlinePos));
                var end = m1.Transform(new Math.Vector2(cell.textLength, underlinePos));
                commands.Add($"{underlineWidth.ToPdfString()} w");
                commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                commands.Add($"S");
            }
            if (cell.FontData.Strike)
            {
                var strikePos = font.fontData.Os2Table.yStrikeoutPosition * scale;
                var strikeWidth = font.fontData.Os2Table.yStrikeoutSize * scale;
                var start = m1.Transform(new Math.Vector2(0, strikePos));
                var end = m1.Transform(new Math.Vector2(cell.textLength, strikePos));
                commands.Add($"{strikeWidth.ToPdfString()} w");
                commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                commands.Add($"S");
            }
            commands.Add("Q");
        }

        public void AddInnerGridLines(PdfTransform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;

            commands.Add("q");
            commands.Add($"{GridLine.Width.ToPdfString()} w");
            commands.Add(PdfColor.Black.ToFillCommand());
            foreach (var line in pl.GridLines)
            {
                string w, h;
                if (line.X1 == line.X2)
                {
                    w = GridLine.Width.ToPdfString();
                    h = System.Math.Abs(line.Y2 - line.Y1).ToPdfString();
                }
                else
                {
                    w = System.Math.Abs(line.X2 - line.X1).ToPdfString();
                    h = GridLine.Width.ToPdfString();
                }
                commands.Add($"{(line.X1).ToPdfString()} {(line.Y1).ToPdfString()} {w} {h} re");
            }
            commands.Add("f");
            commands.Add("Q");
        }

        public void AddOuterGridBorder(PdfTransform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;

            commands.Add("q");
            commands.Add("1.0 w");
            commands.Add("2 J");
            commands.Add("[] 0 d");
            commands.Add(PdfColor.Black.ToStrokeCommand());
            foreach (var line in pl.BorderLines)
            {
                commands.Add($"{line.X1.ToPdfString()} {line.Y1.ToPdfString()} m");
                commands.Add($"{line.X2.ToPdfString()} {line.Y2.ToPdfString()} l");
            }
            commands.Add("S");
            commands.Add("Q");
        }

        public void AddMarginClipping(PdfTransform pageLayout, PdfContentBounds bounds)
        {
            if(pageLayout is not PdfPageLayout pl) return;

            double y = bounds.Top;
            double width = 0d;
            foreach (var line in pl.BorderLines)
            {
                width = System.Math.Max(width, System.Math.Max(line.X1, line.X2));
                y = System.Math.Min(y, System.Math.Min(line.Y1, line.Y2));
            }
            var heightAdjust = y - bounds.Bottom;
            commands.Add($"{bounds.X.ToPdfString()} {(y).ToPdfString()} {(width - bounds.Left).ToPdfString()} {(bounds.Height - heightAdjust).ToPdfString()} re W n");
        }

        internal override string RenderDictionary()
        {
            var content = string.Join("\n", commands.ToArray()) + "\n";
            var bytes = Encoding.ASCII.GetBytes(content);
            return $"<< /Length {bytes.Length} >>\n" + $"stream\n{content}endstream";
        }

        private string FixEscapeCharacters(string text)
        {
            return text.Replace(@"\", @"\\").Replace(@"(", @"\(").Replace(@")", @"\)");
        }
    }
}
