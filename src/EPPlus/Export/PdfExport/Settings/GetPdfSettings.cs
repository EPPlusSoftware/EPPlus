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
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Settings.PdfPageSizes;
using EPPlus.Graphics.Units;
using System;

namespace OfficeOpenXml.Export.PdfExport.Settings
{
    internal static class GetPdfSettings
    {
        internal static PdfPageSettings GetPdfSettingsFromPrinterSettings(ExcelWorkbook workbook, ExcelPrinterSettings eps)
        {
            var settings = new PdfPageSettings(workbook.RenderContext.FontEngine);
            ApplyPrinterSettings(settings, eps);
            return settings;
        }

        internal static PdfPageSettings GetPdfSettingsForSheet(
    PdfPageSettings baseSettings, ExcelPrinterSettings eps)
        {
            var s = baseSettings.CloneForSheet();
            ApplyPrinterSettings(s, eps);
            return s;
        }

        private static void ApplyPrinterSettings(PdfPageSettings settings, ExcelPrinterSettings eps)
        {

            var leftMargin = UnitConversion.ToMillimeters(eps.LeftMargin);
            var rightMargin = UnitConversion.ToMillimeters(eps.RightMargin);
            var topMargin = UnitConversion.ToMillimeters(eps.TopMargin);
            var bottomMargin = UnitConversion.ToMillimeters(eps.BottomMargin);
            var headerMargin = UnitConversion.ToMillimeters(eps.HeaderMargin);
            var footerMargin = UnitConversion.ToMillimeters(eps.FooterMargin);
            settings.Margins = new PdfMargins(topMargin, bottomMargin, leftMargin, rightMargin, headerMargin, footerMargin);
            settings.Orientation = (Orientations)eps.Orientation;
            //Scaling is not yet implemented.
            settings.Scaling = new PdfScaling(eps.Scale);
            settings.ShowHeadings = eps.ShowHeaders;
            settings.RowsToRepeatAtTop = eps.RepeatRows != null ? eps.RepeatRows.Address : null;
            settings.ColumnsToRepeatAtLeft = eps.RepeatColumns != null ? eps.RepeatColumns.Address : null;
            //Print area is implemented and uses the defined name instead of this setting. this setting should override the print area defined name.
            settings.PrintArea = eps.PrintArea != null ? eps.PrintArea.Address : null;
            settings.ShowGridLines = eps.ShowGridLines;
            //Centering is not implemented.
            settings.CenterOnPageHorizontally = eps.HorizontalCentered;
            settings.CenterOnPageVertically = eps.VerticalCentered;
            settings.PageOrders = (PageOrders)eps.PageOrder;
            //Black and white is not yet implemented.
            settings.BlackAndWhite = eps.BlackAndWhite;
            //Draft is not implemtened.
            settings.Draft = eps.Draft;
            settings.PageSize = GetPageSize(eps.PaperSize);
            settings.CommentsAndNotes = (CommentsAndNotes)eps.CellComments;
            settings.CellErrors = (CellErrors)eps.Errors;
            settings.FirstPageNumber = eps.FirstPageNumber;
        }

        private static PdfPageSize GetPageSize(ePaperSize PaperSize)
        {
            switch (PaperSize)
            {
                case ePaperSize.Letter:
                    return PdfPageSize.Letter;
                case ePaperSize.LetterSmall:
                    return PdfPageSize.LetterSmall;
                case ePaperSize.Tabloid:
                    return PdfPageSize.Tabloid;
                case ePaperSize.Ledger:
                    return PdfPageSize.Ledger;
                case ePaperSize.Legal:
                    return PdfPageSize.Legal;
                case ePaperSize.Statement:
                    return PdfPageSize.Statement;
                case ePaperSize.Executive:
                    return PdfPageSize.Executive;
                case ePaperSize.A3:
                    return PdfPageSize.A3;
                case ePaperSize.A4:
                    return PdfPageSize.A4;
                case ePaperSize.A4Small:
                    return PdfPageSize.A4Small;
                case ePaperSize.A5:
                    return PdfPageSize.A5;
                case ePaperSize.B4:
                    return PdfPageSize.B4;
                case ePaperSize.B5:
                    return PdfPageSize.B5;
                case ePaperSize.Folio:
                    return PdfPageSize.Folio;
                case ePaperSize.Quarto:
                    return PdfPageSize.Quarto;
                case ePaperSize.Standard10_14:
                    return PdfPageSize.Standard10_14;
                case ePaperSize.Standard11_17:
                    return PdfPageSize.Standard11_17;
                case ePaperSize.Note:
                    return PdfPageSize.Note;
                case ePaperSize.Envelope9:
                    return PdfPageSize.Envelope9;
                case ePaperSize.Envelope10:
                    return PdfPageSize.Envelope10;
                case ePaperSize.Envelope11:
                    return PdfPageSize.Envelope11;
                case ePaperSize.Envelope12:
                    return PdfPageSize.Envelope12;
                case ePaperSize.Envelope14:
                    return PdfPageSize.Envelope14;
                case ePaperSize.C:
                    return PdfPageSize.C;
                case ePaperSize.D:
                    return PdfPageSize.D;
                case ePaperSize.E:
                    return PdfPageSize.E;
                case ePaperSize.DLEnvelope:
                    return PdfPageSize.DLEnvelope;
                case ePaperSize.C5Envelope:
                    return PdfPageSize.C5Envelope;
                case ePaperSize.C3Envelope:
                    return PdfPageSize.C3Envelope;
                case ePaperSize.C4Envelope:
                    return PdfPageSize.C4Envelope;
                case ePaperSize.C6Envelope:
                    return PdfPageSize.C6Envelope;
                case ePaperSize.C65Envelope:
                    return PdfPageSize.C65Envelope;
                case ePaperSize.B4Envelope:
                    return PdfPageSize.B4Envelope;
                case ePaperSize.B5Envelope:
                    return PdfPageSize.B5Envelope;
                case ePaperSize.B6Envelope:
                    return PdfPageSize.B6Envelope;
                case ePaperSize.ItalyEnvelope:
                    return PdfPageSize.ItalyEnvelope;
                case ePaperSize.MonarchEnvelope:
                    return PdfPageSize.MonarchEnvelope;
                case ePaperSize.Six3_4Envelope:
                    return PdfPageSize.Six3_4Envelope;
                case ePaperSize.USStandard:
                    return PdfPageSize.USStandard;
                case ePaperSize.GermanStandard:
                    return PdfPageSize.GermanStandard;
                case ePaperSize.GermanLegal:
                    return PdfPageSize.GermanLegal;
                case ePaperSize.ISOB4:
                    return PdfPageSize.ISOB4;
                case ePaperSize.JapaneseDoublePostcard:
                    return PdfPageSize.JapaneseDoublePostcard;
                case ePaperSize.Standard9:
                    return PdfPageSize.Standard9;
                case ePaperSize.Standard10:
                    return PdfPageSize.Standard10;
                case ePaperSize.Standard15:
                    return PdfPageSize.Standard15;
                case ePaperSize.InviteEnvelope:
                    return PdfPageSize.InviteEnvelope;
                case ePaperSize.LetterExtra:
                    return PdfPageSize.LetterExtra;
                case ePaperSize.LegalExtra:
                    return PdfPageSize.LegalExtra;
                case ePaperSize.TabloidExtra:
                    return PdfPageSize.TabloidExtra;
                case ePaperSize.A4Extra:
                    return PdfPageSize.A4Extra;
                case ePaperSize.LetterTransverse:
                    return PdfPageSize.LetterTransverse;
                case ePaperSize.A4Transverse:
                    return PdfPageSize.A4Transverse;
                case ePaperSize.LetterExtraTransverse:
                    return PdfPageSize.LetterExtraTransverse;
                case ePaperSize.SuperA:
                    return PdfPageSize.SuperA;
                case ePaperSize.SuperB:
                    return PdfPageSize.SuperB;
                case ePaperSize.LetterPlus:
                    return PdfPageSize.LetterPlus;
                case ePaperSize.A4Plus:
                    return PdfPageSize.A4Plus;
                case ePaperSize.A5Transverse:
                    return PdfPageSize.A5Transverse;
                case ePaperSize.JISB5Transverse:
                    return PdfPageSize.JISB5Transverse;
                case ePaperSize.A3Extra:
                    return PdfPageSize.A3Extra;
                case ePaperSize.A5Extra:
                    return PdfPageSize.A5Extra;
                case ePaperSize.ISOB5:
                    return PdfPageSize.ISOB5;
                case ePaperSize.A2:
                    return PdfPageSize.A2;
                case ePaperSize.A3Transverse:
                    return PdfPageSize.A3Transverse;
                case ePaperSize.A3ExtraTransverse:
                    return PdfPageSize.A3ExtraTransverse;
                default:
                    //should return custom paper size instead.
                    throw new ArgumentOutOfRangeException(nameof(PaperSize), PaperSize, "Unsupported paper size.");
            }

        }
    }
}
