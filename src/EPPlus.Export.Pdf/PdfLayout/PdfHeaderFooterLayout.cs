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
using EPPlus.Export.Pdf.Pdfhelpers;
using OfficeOpenXml;
using OfficeOpenXml.Style.HeaderFooterTextFormat;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace EPPlus.Export.Pdf.PdfLayout
{
    internal class PdfHeaderFooterLayout : PdfTransform
    {
        public PdfCellTextLine textLine = new PdfCellTextLine();

        public PdfHeaderFooterLayout(ExcelHeaderFooterTextCollection textCollection, ExcelWorksheet ws)
        {
            foreach (var text in textCollection)
            {
                switch (text.Text)
                {
                    case "&A":
                        break;
                    case "&D":
                        break;
                    case "&F":
                        break;
                    case "&N":
                        break;
                    case "&P":
                        break;
                    case "&T":
                        break;
                    case "&Z":
                        break;
                    default:
                        break;

                }
                PdfCellTextItem textItem = new PdfCellTextItem();

            }
        }

        public static string GetWorksheetTabName()
        {
        }

        public static string GetDate()
        {
        }

        public static string GetWorkbookFileName()
        {
        }

        public static string GetNumberOfPages()
        {
        }

        public static string GetCurrentPageNumber()
        {
        }

        public static string GetCurrentTime()
        {
        }

        public static string GetWorkbookfilePath()
        {
        }
    }
}
