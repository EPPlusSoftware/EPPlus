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
using EPPlus.Fonts.OpenType;
using EPPlus.Fonts.OpenType.TextShaping;
using OfficeOpenXml.Export.PdfExport.TextMapping;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style.XmlAccess;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.Data
{
    internal class PdfWorksheet
    {
        public Dictionary<string, PdfCommentsAndNotes> CommentsAndNotesCollections = new Dictionary<string, PdfCommentsAndNotes>();
        public List<PdfDrawing> Drawings = new List<PdfDrawing>();
        public List<PdfRange> Ranges = null; //Rename this
        public PdfRange CommentsAndNotes;
        public PdfHeaderFooterCollection HeaderFooters = null;
        public double ZeroCharWidth;
        public int ToRow;
        public int PrintTitleRowFrom = -1;
        public int PrintTitleRowTo = -1;
        public int PrintTitleColFrom = -1;
        public int PrintTitleColTo = -1;

        //EPPlus references
        public ExcelWorksheet Worksheet { get; set; }
        public ExcelWorksheet CommentsAndNotesSheet { get; set; }
        public ExcelNamedStyleXml NormalStyle { get { return Worksheet.Workbook.Styles.GetNormalStyle(); } }
        public FontSubFamily GetSubFamilyFromNormalStyle //move this to a helper class or something.
        {
            get
            {
                var nsf = NormalStyle.Style.Font;
                var SubFamily = FontSubFamily.Regular;
                if (nsf.Bold)
                {
                    SubFamily = FontSubFamily.Bold;
                    if (nsf.Italic)
                    {
                        SubFamily = FontSubFamily.BoldItalic;
                    }
                }
                else if (nsf.Italic)
                {
                    SubFamily = FontSubFamily.Italic;
                }
                return SubFamily;
            }
        }

        public static double GetThemeFont0Width(ExcelWorksheet ws)
        {
            var ns = ws.Workbook.Styles.GetNormalStyle();
            TextShaper shaper = OpenTypeFonts.GetTextShaper(ns.Style.Font.Name, FontSubFamily.Regular);
            var shapedText = shaper.ShapeLight("0");
            return shapedText.GetWidthInPoints(ns.Style.Font.Size);
        }
    }
}