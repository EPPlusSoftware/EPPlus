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
using EPPlus.Export.Pdf.Layout;
using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal static class PdfBookmarkCollector
    {
        internal static void CollectNamedDestinations(PdfWorksheet[] pdfSheets, List<PdfNamedDest> dests)
        {
            if (pdfSheets == null || pdfSheets.Length == 0) return;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            void Add(ExcelNamedRange n)
            {
                if (n == null || string.IsNullOrEmpty(n.Name)) return;
                if (n.Name.StartsWith("_xlnm.")) return;
                if (!string.IsNullOrEmpty(n.NameFormula)) return;
                var targetSheet = n.Worksheet ?? n.LocalSheet;
                if (targetSheet == null || n.Start.Row <= 0 || n.Start.Column <= 0) return;
                if (!seen.Add(targetSheet.Name + "\u0000" + n.Name)) return;
                string cellText = null;
                try
                {
                    cellText = targetSheet.Cells[n.Start.Row, n.Start.Column].Text;
                }
                catch
                {
                }
                dests.Add(new PdfNamedDest
                {
                    Name = n.Name,
                    Sheet = targetSheet.Name,
                    Row = n.Start.Row,
                    Col = n.Start.Column,
                    CellText = cellText
                });
            }
            var wb = pdfSheets[0].Worksheet.Workbook;
            foreach (var n in wb.Names)
            {
                Add(n);
            }
            foreach (var ps in pdfSheets)
            {
                foreach (var n in ps.Worksheet.Names)
                {
                    Add(n);
                }
            }
        }
    }
}
