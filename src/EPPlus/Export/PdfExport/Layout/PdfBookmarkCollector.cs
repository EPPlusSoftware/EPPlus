using EPPlus.Export.Pdf.Layout;
using OfficeOpenXml.Export.PdfExport.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport.Layout
{
    internal class PdfBookmarkCollector
    {
        internal static void CollectNamedDestinations(PdfWorksheet[] pdfSheets, List<PdfNamedDest> dests)
        {
            if (pdfSheets == null || pdfSheets.Length == 0) return;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void Add(ExcelNamedRange n)
            {
                if (n == null || string.IsNullOrEmpty(n.Name)) return;
                if (n.Name.StartsWith("_xlnm.")) return;            // built-in print area / titles
                if (!string.IsNullOrEmpty(n.NameFormula)) return;   // formula-only name, no cell target

                // Sheet the name's cell lives on: workbook-scoped names expose it via .Worksheet;
                // a sheet-scoped name with an unqualified RefersTo has a null .Worksheet, so fall
                // back to its scope sheet (LocalSheet).
                var targetSheet = n.Worksheet ?? n.LocalSheet;
                if (targetSheet == null || n.Start.Row <= 0 || n.Start.Column <= 0) return;

                // Dedup per (sheet, name): the same bm_ name may legitimately recur on several
                // sheets (one section each). A global dedup drops all but the first and leaves
                // later sheets with no sub-bookmarks.
                if (!seen.Add(targetSheet.Name + "\u0000" + n.Name)) return;

                string cellText = null;
                try { cellText = targetSheet.Cells[n.Start.Row, n.Start.Column].Text; }
                catch { /* leave null; writer falls back to the name */ }

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
            foreach (var n in wb.Names) Add(n);                     // workbook-scoped first
            foreach (var ps in pdfSheets)
                foreach (var n in ps.Worksheet.Names) Add(n);       // then sheet-scoped
        }
    }
}
