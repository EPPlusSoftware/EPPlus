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
namespace OfficeOpenXml.Export.PdfExport
{
    /// <summary>
    /// Document-level options for exporting a workbook, worksheets or ranges to PDF.
    /// <para>
    /// These are settings that have no representation in the worksheet's print setup
    /// (<see cref="ExcelPrinterSettings"/>). Page layout — margins, orientation, paper size,
    /// gridlines, headings, repeat rows/columns, print area, comments, errors and so on — is
    /// not configured here; it is taken from each worksheet's <see cref="ExcelPrinterSettings"/>
    /// at export time. Where a setting below can also be varied per worksheet, the worksheet-level
    /// override takes precedence over the value set here.
    /// </para>
    /// </summary>
    public class PdfExportSettings
    {
        /// <summary>
        /// THIS FEATURE IS NOT SUPPORTED.
        /// If true, the subsetted fonts used in the document are embedded into the PDF.
        /// Embedding produces a self-contained, portable file at the cost of size. Default is <c>true</c>.
        /// </summary>
        private bool EmbeddFonts { get; set; } = true;

        /// <summary>
        /// Prefix that marks a defined name as a PDF bookmark. The prefix is stripped from the
        /// bookmark title. The bookmark text is taken from the target cell's content, or the defined
        /// name (minus the prefix) when the cell is empty. Default is <c>"bm_"</c>.
        /// This is the document-wide default; it can be overridden per worksheet.
        /// </summary>
        public string BookmarkNamePrefix { get; set; } = "bm_";

        /// <summary>
        /// If true, the export adds a single hidden, read-only form field so the document presents
        /// itself as an already-prepared form. This stops Adobe Acrobat / Reader from auto-detecting
        /// the cell grid as fillable form fields. Default is <c>true</c>.
        /// </summary>
        public bool SuppressFormFieldDetection { get; set; } = true;

        /// <summary>
        /// Value written to the PDF's /Author metadata. When null or empty, the workbook's
        /// document properties are used.
        /// </summary>
        public string Author { get; set; }

        /// <summary>
        /// Value written to the PDF's /Title metadata. When null or empty, the workbook's
        /// document properties are used.
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// The compression level applied to PDF content streams and images. Trades export
        /// speed against file size. Default is <see cref="PdfCompressionLevel.Smallest"/>.
        /// </summary>
        public PdfCompressionLevel CompressionLevel { get; set; } = PdfCompressionLevel.Smallest;
    }
}