/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf.Helpers;
using System.Drawing;
using System.Text;

namespace EPPlus.Export.Pdf.DocumentObjects.Patterns
{
    /// <summary>
    /// Renders a cell fill pattern from an 8x8 <see cref="ExcelPatternMask"/>.
    /// Replaces the former per-pattern PdfPatternFill subclasses: the geometry now
    /// comes from a single verified mask catalog (taken from Excel's own PDF output)
    /// instead of hand-coded rectangle coordinates.
    ///
    /// The whole 8x8 tile is filled with the Background color, then a rectangle is
    /// drawn with the Foreground color for every foreground cell of the mask.
    ///
    /// The mask stores row 0 as the TOP row, while a PDF content stream has its
    /// origin at the bottom-left with y increasing upward. The mask row r is
    /// therefore emitted at PDF y = 7 - r (a mirror in y). Horizontally adjacent
    /// foreground cells on the same row are merged into a single wider rectangle
    /// to keep the content stream small.
    /// </summary>
    internal class PdfPatternMaskFill : PdfPatternFill
    {
        private readonly byte[,] _mask;

        public PdfPatternMaskFill(ExcelPatternMask pattern, Color foreground, Color background)
            : base(foreground, background)
        {
            _mask = ExcelPatternMaskData.GetMask(pattern);
        }

        public override string CreatePatternResource()
        {
            const int size = 8;
            var sb = new StringBuilder();

            // Fill the whole tile with the background color.
            sb.Append(Background.ToFillCommand());
            sb.Append("\n0 0 ");
            sb.Append(size.ToString());
            sb.Append(" ");
            sb.Append(size.ToString());
            sb.Append(" re\nf\n");

            // Draw foreground rectangles. mask value 0 == foreground.
            sb.Append(Foreground.ToFillCommand());
            sb.Append("\n");
            for (int row = 0; row < size; row++)
            {
                int pdfY = (size - 1) - row; // mirror in y
                int col = 0;
                while (col < size)
                {
                    if (_mask[row, col] == 0)
                    {
                        int start = col;
                        while (col < size && _mask[row, col] == 0)
                        {
                            col++;
                        }
                        int width = col - start;
                        sb.Append(start.ToString());
                        sb.Append(" ");
                        sb.Append(pdfY.ToString());
                        sb.Append(" ");
                        sb.Append(width.ToString());
                        sb.Append(" 1 re\n");
                    }
                    else
                    {
                        col++;
                    }
                }
            }
            sb.Append("f");
            return sb.ToString();
        }
    }
}