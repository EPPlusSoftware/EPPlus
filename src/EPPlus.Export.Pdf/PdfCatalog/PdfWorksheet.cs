using EPPlus.Export.Pdf.PdfLayout;
using EPPlus.Fonts.OpenType;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Interfaces.Fonts;
using OfficeOpenXml.Style.XmlAccess;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfCatalog
{
    internal class PdfWorksheet
    {
        public Dictionary<string, PdfCommentsAndNotes> CommentsAndNotesCollections = new Dictionary<string, PdfCommentsAndNotes>();

        public List<PdfRange> Ranges = null; //Rename this
        public PdfRange CommentsAndNotes;
        public PdfHeaderFooterCollection HeaderFooters = null;
        public double ZeroCharWidth;
        public int ToRow;

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
            FontMeasurerTrueType fontMeasurerTrueType = new FontMeasurerTrueType();
            MeasurementFont font = new MeasurementFont();
            var ns = ws.Workbook.Styles.GetNormalStyle();
            font.FontFamily = ns.Style.Font.Name;
            font.Size = ns.Style.Font.Size;
            font.Style = MeasurementFontStyles.Regular;
            var result = fontMeasurerTrueType.MeasureText("0", font);
            return result.Width;
        }
    }
}