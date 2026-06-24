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
using EPPlus.Graphics;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Export.Pdf.Resources;
using OfficeOpenXml.Export.PdfExport.Data;
using OfficeOpenXml.Export.PdfExport.Layout;
using OfficeOpenXml.Export.PdfExport.RowResize;
using OfficeOpenXml.Export.PdfExport.TextMapping;
using OfficeOpenXml.Export.PdfExport.TextShaping;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace OfficeOpenXml.Export.PdfExport
{
    internal class PdfCatalog
    {
        internal PdfDictionaries _dictionaries = new PdfDictionaries();
        private bool _addTextForHeadings = true;

        //
        // CONSTRUCTORS FOR MULTIPLE WORKSHEETS AS INPUT
        //
        public PdfCatalog() { }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorkbook workbook)
        {
            HandleWorksheetCollection(pageSettings, workbook.Worksheets.ToArray());
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            HandleWorksheetCollection(pageSettings, worksheets);
        }

        public PdfCatalog(PdfPageSettings pageSettings, List<ExcelWorksheet> worksheets)
        {
            HandleWorksheetCollection(pageSettings, worksheets.ToArray());
        }

        private void HandleWorksheetCollection(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            var pdfSheets = GetPdfWorksheets(pageSettings, worksheets);
            foreach (var pdfSheet in pdfSheets)
            {
                ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
                PdfCalculateRowHeight.ResizeRowHeights(pdfSheet);
            }
            var Layout = GetLayout(pageSettings, pdfSheets);
            //send Layout to pdf export here.
        }

        //
        // CONSTRUCTORS FOR SINGLE WORKSHEET AS INPUT
        //

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet worksheet, string fileName)
        {
            BuildPdf(pageSettings, worksheet, (excelPdf, layout) =>
                excelPdf.CreatePdf(pageSettings, _dictionaries, layout, fileName));
        }

        public PdfCatalog(PdfPageSettings pageSettings, ExcelWorksheet worksheet, Stream stream)
        {
            BuildPdf(pageSettings, worksheet, (excelPdf, layout) =>
                excelPdf.CreatePdf(pageSettings, _dictionaries, layout, stream));
        }

        private void BuildPdf(PdfPageSettings pageSettings, ExcelWorksheet worksheet, Action<ExcelPdf, Transform> writePdf)
        {
            //pageSettings.defaultFontName = worksheet.Workbook.ThemeManager.CurrentTheme.FontScheme.MinorFont[0].Typeface;
            pageSettings.defaultFontName = worksheet.Workbook.ThemeManager.GetOrCreateTheme().FontScheme.MinorFont[0].Typeface;
            PdfWorksheet pdfSheet = null;
            try
            {
                Stopwatch sw = Stopwatch.StartNew();

                //Collect Text
                pdfSheet = GetPdfWorksheet(pageSettings, worksheet);
                sw.Stop();
                var CollectTextTime = sw.ElapsedMilliseconds;
                sw.Reset();
                sw.Start();

                //Shape Text
                ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
                sw.Stop();
                var ShapeTextTime = sw.ElapsedMilliseconds;
                sw.Reset();
                sw.Start();

                //Auto-Fit Rows
                PdfCalculateRowHeight.ResizeRowHeights(pdfSheet);
                sw.Stop();
                var AutoFitRowTime = sw.ElapsedMilliseconds;
                sw.Reset();
                sw.Start();

                //Create Layout
                var Layout = GetLayout(pageSettings, pdfSheet);
                sw.Stop();
                var CreateLayoutTime = sw.ElapsedMilliseconds;
                sw.Reset();
                sw.Start();

                //Create Pdf //Done in pdf export project
                ExcelPdf excelPdf = new ExcelPdf();
                writePdf(excelPdf, Layout);
                sw.Stop();
                var CreatePdfTime = sw.ElapsedMilliseconds;
                sw.Reset();
            }
            finally
            {
                //Clean up the temporary worksheet used to build the comments/notes pages,
                //so the source workbook isn't permanently mutated by the PDF export.
                if (pdfSheet != null && pdfSheet.CommentsAndNotesSheet != null)
                {
                    worksheet.Workbook.Worksheets.Delete(pdfSheet.CommentsAndNotesSheet);
                    pdfSheet.CommentsAndNotesSheet = null;
                }
            }
        }

        //
        // CONSTRUCTORS FOR RANGE AS INPUT
        //

        public PdfCatalog(PdfPageSettings pageSettings, ExcelRangeBase range)
        {
            PdfWorksheet pdfSheet = GetPdfWorksheet(pageSettings, range);
            ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
            PdfCalculateRowHeight.ResizeRowHeights(pdfSheet);
            var Layout = GetLayout(pageSettings, pdfSheet);
            //send Layout to pdf export here.
        }

        internal PdfCellCollection GetCellCollectionFromRange(PdfPageSettings pageSettings, ExcelRangeBase range)
        {
            PdfWorksheet pdfSheet = GetPdfWorksheet(pageSettings, range);
            ShapeTextInPdfWorksheet(pageSettings, pdfSheet);
            return pdfSheet.Ranges[0].Map;
        }


        //Private Methods


        //Create Layout Methods

        private Transform GetLayout(PdfPageSettings pageSettings, PdfWorksheet[] pdfSheets)
        {
            var Layout = PdfLayout.GetLayout(pageSettings, _dictionaries, pdfSheets);
            return Layout;
        }

        private Transform GetLayout(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            PdfWorksheet[] pdfSheets = new PdfWorksheet[1] { pdfSheet };
            var Layout = PdfLayout.GetLayout(pageSettings, _dictionaries, pdfSheets);
            return Layout;
        }

        //Shape Text Methods

        internal void ShapeTextInPdfWorksheet(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            // Pass 1: collect text per font
            IterateCells(pdfSheet, cell => PdfTextShaper.CollectText(_dictionaries, cell));

            // Pass 2: build one provider per font
            foreach (var kvp in _dictionaries.Fonts)
            {
                _dictionaries.ShapedProviders[kvp.Key] = kvp.Value.fontSubsetManager.CreateSubsettedProvider();
            }

            // Pass 3: shape text using the pre-built providers
            IterateCells(pdfSheet, cell => PdfTextShaper.ShapeText(pageSettings, _dictionaries, cell));
        }

        private void IterateCells(PdfWorksheet pdfSheet, System.Action<PdfCell> action)
        {
            foreach (var range in pdfSheet.Ranges)
            {
                for (int i = range.Map.FromRow; i <= range.Map.ToRow; i++)
                {
                    for (int j = range.Map.FromColumn; j <= range.Map.ToColumn; j++)
                    {
                        action(range.Map[i, j]);
                    }
                }
            }

            if (pdfSheet.CommentsAndNotes.Map != null)
            {
                for (int i = pdfSheet.CommentsAndNotes.Map.FromRow; i <= pdfSheet.CommentsAndNotes.Map.ToRow; i++)
                {
                    for (int j = pdfSheet.CommentsAndNotes.Map.FromColumn; j <= pdfSheet.CommentsAndNotes.Map.ToColumn; j++)
                    {
                        action(pdfSheet.CommentsAndNotes.Map[i, j]);
                    }
                }
            }
        }

        //Collect Text Methods
        private PdfWorksheet[] GetPdfWorksheets(PdfPageSettings pageSettings, ExcelWorksheet[] worksheets)
        {
            PdfWorksheet[] pdfSheets = new PdfWorksheet[worksheets.Length];
            for (int i = 0; i < pdfSheets.Length; i++)
            {
                pdfSheets[i] = GetPdfWorksheet(pageSettings, worksheets[i]);
            }
            return pdfSheets;
        }

        internal PdfWorksheet GetPdfWorksheet(PdfPageSettings pageSettings, ExcelWorksheet worksheet)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>();
            pdfSheet.Worksheet = worksheet;
            pdfSheet.Ranges = GetRanges(pdfSheet.Worksheet);
            if (pageSettings.ShowHeadings && _addTextForHeadings) _dictionaries.AddFont(pageSettings, pdfSheet.NormalStyle.Style.Font.Name, pdfSheet.GetSubFamilyFromNormalStyle, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890");
            _addTextForHeadings = false;
            GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges);
            GetPrintTitles(pageSettings, pdfSheet);
            GetHeaderFooter(pageSettings, pdfSheet);
            GetCommentsAndNotes(pageSettings, pdfSheet);
            return pdfSheet;
        }

        private PdfWorksheet GetPdfWorksheet(PdfPageSettings pageSettings, ExcelRangeBase excelRange)
        {
            PdfWorksheet pdfSheet = new PdfWorksheet();
            pdfSheet.Ranges = new List<PdfRange>();
            pdfSheet.Worksheet = excelRange.Worksheet;
            pdfSheet.Ranges.Add(new PdfRange(excelRange, false));
            if (pageSettings.ShowHeadings && _addTextForHeadings) _dictionaries.AddFont(pageSettings, pdfSheet.NormalStyle.Style.Font.Name, pdfSheet.GetSubFamilyFromNormalStyle, "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890");
            _addTextForHeadings = false;
            pdfSheet.Ranges[0] = GetMaps(pageSettings, pdfSheet, pdfSheet.Ranges[0]);
            GetPrintTitles(pageSettings, pdfSheet);
            GetHeaderFooter(pageSettings, pdfSheet);
            GetCommentsAndNotes(pageSettings, pdfSheet);
            return pdfSheet;
        }

        private List<PdfRange> GetRanges(ExcelWorksheet worksheet)
        {
            List<PdfRange> ranges = new List<PdfRange>();
            if (worksheet.Names.ContainsKey("_xlnm.Print_Area"))
            {
                for (int i = 0; i < worksheet.Names["_xlnm.Print_Area"].Addresses.Count; i++)
                {
                    var range = worksheet.Cells[worksheet.Names["_xlnm.Print_Area"].Addresses[i].Address];
                    ranges.Add(new PdfRange(range, false));
                }
            }
            else
            {
                var range = worksheet.DimensionByValue;
                var pdfRange = new PdfRange(range, true);
                pdfRange.ExtendColumns = true;
                ranges.Add(pdfRange);
            }
            return ranges;
        }

        private void GetMaps(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, List<PdfRange> ranges)
        {
            for (int i = 0; i < ranges.Count; i++)
            {
                ranges[i] = GetMaps(pageSettings, pdfSheet, ranges[i]);
            }
        }

        private PdfRange GetMaps(PdfPageSettings pageSettings, PdfWorksheet pdfSheet, PdfRange range)
        {
            var temp = range;
            temp.Map = PdfTextMap.SetTextMap(pageSettings, _dictionaries, pdfSheet, ref temp);
            range = temp;
            return range;
        }

        private void GetPrintTitles(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            var worksheet = pdfSheet.Worksheet;
            // --- Step 1: auto-detect from the worksheet's _xlnm.Print_Titles defined name ---
            if (worksheet.Names.ContainsKey("_xlnm.Print_Titles"))
            {
                var printTitlesName = worksheet.Names["_xlnm.Print_Titles"];
                foreach (var address in printTitlesName.Addresses)
                {
                    // A full-row reference spans every column  (e.g. $1:$3  →  _toCol == MaxColumns)
                    if (address._toCol >= ExcelPackage.MaxColumns)
                    {
                        pdfSheet.PrintTitleRowFrom = address._fromRow;
                        pdfSheet.PrintTitleRowTo = address._toRow;
                    }
                    // A full-column reference spans every row  (e.g. $A:$B  →  _toRow == MaxRows)
                    else if (address._toRow >= ExcelPackage.MaxRows)
                    {
                        pdfSheet.PrintTitleColFrom = address._fromCol;
                        pdfSheet.PrintTitleColTo = address._toCol;
                    }
                }
            }
            // --- Step 2: PdfPageSettings overrides take precedence over the defined name ---
            if (pageSettings.RowsToRepeatAtTop != null)
            {
                ExcelAddressBase repeatRows = new ExcelAddressBase(pageSettings.RowsToRepeatAtTop);
                pdfSheet.PrintTitleRowFrom = repeatRows._fromRow;
                pdfSheet.PrintTitleRowTo = repeatRows._toRow;
            }
            if (pageSettings.ColumnsToRepeatAtLeft != null)
            {
                ExcelAddressBase repeatCols = new ExcelAddressBase(pageSettings.ColumnsToRepeatAtLeft);
                pdfSheet.PrintTitleColFrom = repeatCols._fromCol;
                pdfSheet.PrintTitleColTo = repeatCols._toCol;
            }

            // --- Step 3: mark cells so the renderer can identify them instantly ---
            foreach (var range in pdfSheet.Ranges)
                MarkPrintTitleCells(pdfSheet, range);
        }

        private static void MarkPrintTitleCells(PdfWorksheet pdfSheet, PdfRange range)
        {
            var map = range.Map;

            for (int row = map.FromRow; row <= map.ToRow; row++)
            {
                bool isTitleRow = pdfSheet.PrintTitleRowFrom >= 0
                               && row >= pdfSheet.PrintTitleRowFrom
                               && row <= pdfSheet.PrintTitleRowTo;

                for (int col = map.FromColumn; col <= map.ToColumn; col++)
                {
                    bool isTitleCol = pdfSheet.PrintTitleColFrom >= 0
                                   && col >= pdfSheet.PrintTitleColFrom
                                   && col <= pdfSheet.PrintTitleColTo;

                    if (!isTitleRow && !isTitleCol) continue;

                    var cell = map[row, col];
                    if (cell == null) continue;

                    if (isTitleRow) cell.IsPrintTitleRow = true;
                    if (isTitleCol) cell.IsPrintTitleCol = true;
                }
            }
        }

        private void GetHeaderFooter(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            pdfSheet.HeaderFooters = new PdfHeaderFooterCollection(pageSettings, _dictionaries, pdfSheet, pdfSheet.Worksheet.HeaderFooter);
        }

        private void GetCommentsAndNotes(PdfPageSettings pageSettings, PdfWorksheet pdfSheet)
        {
            if (pageSettings.CommentsAndNotes == CommentsAndNotes.AtEndOfSheet && pdfSheet.CommentsAndNotesCollections.Count > 0)
            {
                var cnPageSettings = new PdfPageSettings();
                cnPageSettings.CommentsAndNotes = CommentsAndNotes.None;
                cnPageSettings.ShowHeadings = false;
                pdfSheet.CommentsAndNotesSheet = PdfCommentsAndNotes.CreateCommentAndNotesPages(pdfSheet.CommentsAndNotesCollections, pdfSheet.Worksheet);
                pdfSheet.CommentsAndNotes = new PdfRange(pdfSheet.CommentsAndNotesSheet.Dimension, false);
                pdfSheet.CommentsAndNotes = GetMaps(cnPageSettings, pdfSheet, pdfSheet.CommentsAndNotes);
            }
        }
    }
}