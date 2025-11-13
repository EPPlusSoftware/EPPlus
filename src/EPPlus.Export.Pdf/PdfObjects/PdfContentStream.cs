/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/
using EPPlus.Export.Pdf.Math;
using EPPlus.Export.Pdf.PdfGraphics;
using EPPlus.Export.Pdf.Pdfhelpers;
using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Export.Pdf.PdfResources;
using EPPlus.Export.Pdf.PdfSettings;
using OfficeOpenXml.Style;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects
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
            if (cell.CellFillData.GradientFillData != null && cell.CellFillData.PattenStyle != ExcelFillStyle.Solid)
            {
                commands.Add($"% Pattern Start: {cell.Name}");
                commands.Add("q");
                commands.Add("/Pattern cs");
                commands.Add($"/{label} scn");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("f");
                commands.Add("Q");
                commands.Add($"% Pattern End: {cell.Name}");
            }
            else if (cell.CellFillData.BackgroundColor != null && cell.CellFillData.PattenStyle == ExcelFillStyle.Solid)
            {
                commands.Add($"% Solid Fill Start: {cell.Name}");
                commands.Add("q");
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add( cell.CellFillData.enhanceGridLine ? PdfColor.Black.ToStrokeCommand() : cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
                commands.Add("Q");
                commands.Add($"% Solid Fill End: {cell.Name}");
            }
            else if (cell.CellFillData.BackgroundColor != null && cell.CellFillData.PattenStyle != ExcelFillStyle.None)
            {
                commands.Add($"% Gradient Start: {cell.Name}");
                commands.Add("q");
                commands.Add($"{GridLine.HalfWidth.ToPdfString()} w");
                commands.Add(cell.CellFillData.BackgroundColor.ToFillCommand());
                commands.Add(cell.CellFillData.BackgroundColor.ToStrokeCommand());
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("B");
                commands.Add("Q");
                commands.Add("q");
                commands.Add("/Pattern cs");
                commands.Add($"/{label} scn");
                commands.Add($"{cell.LocalPosition.X.ToPdfString()} {cell.LocalPosition.Y.ToPdfString()} {cell.Size.X.ToPdfString()} {cell.Size.Y.ToPdfString()} re");
                commands.Add("f");
                commands.Add("Q");
                commands.Add($"% Gradient End: {cell.Name}");
            }
        }

        public void AddBorderLayout(PdfCellBorderLayout cell)
        {
            var borderRenderer = new PdfBorderRenderer(cell);
            borderRenderer.RenderBorder(this);
            //commands.Add($"% Border Start: {cell.Name}");
            //commands.Add("q");
            //borderRenderer.RenderBorder(this, cell.BorderData.Top, LineType.Horizontal, cell.LocalPosition.X, cell.LocalPosition.Y + cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y + cell.Size.Y);
            //borderRenderer.RenderBorder(this, cell.BorderData.Bottom, LineType.Horizontal, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y);
            //borderRenderer.RenderBorder(this, cell.BorderData.Left, LineType.Vertical, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X, cell.LocalPosition.Y + cell.Size.Y);
            //borderRenderer.RenderBorder(this, cell.BorderData.Right, LineType.Vertical, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y + cell.Size.Y);
            //borderRenderer.RenderBorder(this, cell.BorderData.DiagonalUp, LineType.DiagonalUp, cell.LocalPosition.X, cell.LocalPosition.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y + cell.Size.Y);
            //borderRenderer.RenderBorder(this, cell.BorderData.DiagonalDown, LineType.DiagonalDown, cell.LocalPosition.X, cell.LocalPosition.Y + cell.Size.Y, cell.LocalPosition.X + cell.Size.X, cell.LocalPosition.Y);
            //commands.Add("Q");
            //commands.Add($"% Border End: {cell.Name}");
        }


        //Get font label //need to update this one too for same reasons as AddFontData
        internal PdfFontResource GetFontResource(PdfDictionaries Dictionaries, PdfPageSettings PageSettings, string fontName, string subFamily, double fontSize)
        {
            if (!Dictionaries.Fonts.ContainsKey(fontName))
            {
                int label = 1;
                if (Dictionaries.Fonts.Count > 0)
                {
                    label = Dictionaries.Fonts.Last().Value.labelNumber + 1;
                }
                PdfFontResource fr = new PdfFontResource(fontName, subFamily, label, PageSettings);
                if (fontName != "Courier New")
                {
                    //Document.Add(fr.GetFontDescriptorObject(Document.Count + 1));
                    //Document.Add(fr.GetWidthsObject(Document.Count + 1));
                }
                //Document.Add(fr.GetFontObject(Document.Count + 1));
                Dictionaries.Fonts.Add(fontName, fr);
            }
            return Dictionaries.Fonts[fontName];
        }

        public void AddCellContentLayout(PdfCellContentLayout cell, PdfDictionaries dictionaries, PdfPageSettings pageSettings)
        {
            commands.Add($"% Content Start: {cell.Name}");
            commands.Add("q");
            if (cell.Clip)
            {
                commands.Add($"{cell.Clipping.X.ToPdfString()} {cell.Clipping.Y.ToPdfString()} {cell.Clipping.Width.ToPdfString()} {cell.Clipping.Height.ToPdfString()} re W n");
            }
            double rot = cell.CellAlignmentData.TextRotation * System.Math.PI / 180.0;
            double lineLength = 0;
            Matrix3x3 lineMatrix = new Matrix3x3(System.Math.Cos(rot), System.Math.Sin(rot), -System.Math.Sin(rot), System.Math.Cos(rot), cell.LocalPosition.X, cell.LocalPosition.Y);
            for (int j = 0; j < cell.TextLines.Count; j++)
            {
                var Line = cell.TextLines[j];
                Matrix3x3 textRunMatrix = lineMatrix;
                Matrix3x3 modifierMatrix = Matrix3x3.Identity;
                bool useModifiedMatrix = false;
                commands.Add("BT");
                commands.Add($"{lineMatrix.A.ToPdfString()} {lineMatrix.B.ToPdfString()} {lineMatrix.C.ToPdfString()} {lineMatrix.D.ToPdfString()} {lineMatrix.E.ToPdfString()} {lineMatrix.F.ToPdfString()} Tm");
                for (int i=0; i<Line.TextItemCollection.Count; i++)
                {
                    var fontData = Line.TextItemCollection[i];
                    var font = GetFontResource(dictionaries, pageSettings, fontData.FullFontName, fontData.SubFamily, fontData.FontSize);
                    double size = fontData.FontSize;
                    double scale = fontData.FontSize / font.fontData.HeadTable.UnitsPerEm;
                    if (fontData.Bold)
                    {
                        commands.Add("0.25 w");
                        commands.Add("2 Tr");
                        commands.Add(fontData.FontColor.ToStrokeCommand());
                    }
                    else
                    {
                        commands.Add("0 Tr");
                    }
                    if (fontData.Italic)
                    {
                        var ia = font.fontData.PostTable.italicAngle;
                        if (ia <= 0) ia = 12d * System.Math.PI / 180.0d;
                        modifierMatrix.C = System.Math.Tan(ia);
                        modifierMatrix = modifierMatrix * textRunMatrix;
                        useModifiedMatrix = true;
                    }
                    if (fontData.SuperScript)
                    {
                        var supOffX = font.fontData.Os2Table.ySuperscriptXOffset * scale;
                        var supOffY = font.fontData.Os2Table.ySuperscriptYOffset * scale;
                        var supSizeX = font.fontData.Os2Table.ySuperscriptXSize * scale;
                        var supSizeY = font.fontData.Os2Table.ySuperscriptYSize * scale;
                        modifierMatrix.E = lineMatrix.E + supOffX;
                        modifierMatrix.F = lineMatrix.F + supOffY;
                        size = supSizeY;
                        useModifiedMatrix = true;
                    }
                    else if (fontData.SubScript)
                    {
                        var supOffX = font.fontData.Os2Table.ySubscriptXOffset * scale;
                        var supOffY = font.fontData.Os2Table.ySubscriptYOffset * scale;
                        var supSizeX = font.fontData.Os2Table.ySubscriptXSize * scale;
                        var supSizeY = font.fontData.Os2Table.ySubscriptYSize * scale;
                        modifierMatrix.E = lineMatrix.E + supOffX;
                        modifierMatrix.F = lineMatrix.F + supOffY;
                        size = supSizeY;
                        useModifiedMatrix = true;
                    }
                    if (useModifiedMatrix) commands.Add($"{modifierMatrix.A.ToPdfString()} {modifierMatrix.B.ToPdfString()} {modifierMatrix.C.ToPdfString()} {modifierMatrix.D.ToPdfString()} {modifierMatrix.E.ToPdfString()} {modifierMatrix.F.ToPdfString()} Tm");
                    commands.Add($"/{font.Label} {size.ToPdfString()} Tf");
                    commands.Add(fontData.FontColor.ToFillCommand());
                    commands.Add($"({FixEscapeCharacters(fontData.Text)}) Tj");
                    if (fontData.Underline)
                    {
                        var underlinePos = font.fontData.PostTable.underlinePosition * scale;
                        var underlineWidth = font.fontData.PostTable.underlineThickness * scale;
                        var start = textRunMatrix.Transform(new Vector2(0, underlinePos));
                        var end = textRunMatrix.Transform(new Vector2(fontData.TextLength, underlinePos));
                        commands.Add($"{underlineWidth.ToPdfString()} w");
                        commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                        commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                        commands.Add($"S");
                    }
                    if (fontData.Strike)
                    {
                        var strikePos = font.fontData.Os2Table.yStrikeoutPosition * scale;
                        var strikeWidth = font.fontData.Os2Table.yStrikeoutSize * scale;
                        var start = textRunMatrix.Transform(new Vector2(0, strikePos));
                        var end = textRunMatrix.Transform(new Vector2(fontData.TextLength, strikePos));
                        commands.Add($"{strikeWidth.ToPdfString()} w");
                        commands.Add($"{start.X.ToPdfString()} {start.Y.ToPdfString()} m");
                        commands.Add($"{end.X.ToPdfString()} {end.Y.ToPdfString()} l");
                        commands.Add($"S");
                    }
                    lineLength += fontData.TextLength;
                    textRunMatrix = textRunMatrix * Matrix3x3.Translation(fontData.TextLength, 0);
                    if (useModifiedMatrix) commands.Add($"{textRunMatrix.A.ToPdfString()} {textRunMatrix.B.ToPdfString()} {textRunMatrix.C.ToPdfString()} {textRunMatrix.D.ToPdfString()} {textRunMatrix.E.ToPdfString()} {textRunMatrix.F.ToPdfString()} Tm");
                    useModifiedMatrix = false;
                }
                lineMatrix = textRunMatrix * Matrix3x3.Translation(-lineLength, -Line.LineHeight);
                commands.Add("ET");
            }
            commands.Add("Q");
            commands.Add($"% Content End: {cell.Name}");
        }

        public void AddInnerGridLines(PdfTransform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;

            commands.Add($"% Gridlnes Start");
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
                commands.Add($"{line.X1.ToPdfString()} {line.Y1.ToPdfString()} {w} {h} re");
            }
            commands.Add("f");
            commands.Add("Q");
            commands.Add($"% Gridlines End");
        }

        public void AddOuterGridBorder(PdfTransform pageLayout)
        {
            if (pageLayout is not PdfPageLayout pl) return;

            commands.Add($"% Gridlines Border Start");
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
            commands.Add($"% Gridlines Border End");
        }

        public void AddMarginClipping(PdfTransform pageLayout, PdfContentBounds bounds)
        {
            if(pageLayout is not PdfPageLayout pl) return;

            commands.Add($"% Margin Clip Start");
            double y = bounds.Top;
            double width = 0d;
            foreach (var line in pl.BorderLines)
            {
                width = System.Math.Max(width, System.Math.Max(line.X1, line.X2));
                y = System.Math.Min(y, System.Math.Min(line.Y1, line.Y2));
            }
            var heightAdjust = y - bounds.Bottom;
            commands.Add($"{bounds.X.ToPdfString()} {y.ToPdfString()} {(width - bounds.Left).ToPdfString()} {(bounds.Height - heightAdjust).ToPdfString()} re W n");
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
