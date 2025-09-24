using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.Pdfhelpers;
using OfficeOpenXml.PDF.PdfLayout;
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

        public void AddCellLayout(PdfCellLayout cell, string PatternLabel)
        {
            if (cell.CellFillData.GradientFillData != null)
            {
                commands.Add("q");
                commands.Add("/Pattern cs");
                commands.Add($"/{PatternLabel} scn");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("f");
                commands.Add("Q");
            }
            else if (cell.CellFillData.BackgroundColor != null && cell.CellFillData.BackgroundColor.A >= 0.99999f)
            {
                commands.Add("q");
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add(cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {(cell.LocalPosition.Y).ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
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

        public void AddCellContentLayout(PdfCellContentLayout cell, string fontLabel)
        {
            commands.Add("q");
            if (cell.Clip)
            {
                commands.Add($"{cell.Clipping.X.ToPdfString()} {cell.Clipping.Y.ToPdfString()} {cell.Clipping.Width.ToPdfString()} {cell.Clipping.Height.ToPdfString()} re W n");
            }
            commands.Add("BT");
            commands.Add($"/{fontLabel} {cell.FontData.FontSize.ToPdfString()} Tf");
            commands.Add(cell.FontData.FontColor.ToFillCommand());
            double rot = cell.CellAlignmentData.TextRotation * System.Math.PI / 180.0;
            commands.Add($"{System.Math.Cos(rot).ToPdfString()} {System.Math.Sin(rot).ToPdfString()} {(-System.Math.Sin(rot)).ToPdfString()} {System.Math.Cos(rot).ToPdfString()} {cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} Tm");
            commands.Add($"({FixEscapeCharacters(cell.FontData.Text)}) Tj");
            commands.Add("ET");
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
