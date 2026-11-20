using EPPlus.Fonts.OpenType;
using OfficeOpenXml.Interfaces.Fonts;
using System.Collections.Generic;

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
        /// If true, the subsetted fonts used in the document are embedded into the PDF.
        /// Embedding produces a self-contained, portable file at the cost of size. Default is <c>true</c>.
        /// </summary>
        public bool EmbeddFonts { get; set; } = true;

        /// <summary>
        /// If true, EPPlus also looks for fonts in the operating system's font directories
        /// (for example c:\windows\fonts). Default is <c>true</c>.
        /// </summary>
        public bool SearchSystemDirectories { get; set; } = true;

        /// <summary>
        /// Additional folders to search for fonts, beyond the system directories.
        /// </summary>
        public List<string> FontDirectories { get; } = new List<string>();

        /// <summary>
        /// Font family used when a cell's font cannot be resolved. Empty means EPPlus picks a default.
        /// </summary>
        public string DefaultFontName { get; set; } = "";

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
        /// GSUB (glyph substitution) features requested when shaping text, e.g. ligatures and
        /// contextual alternates. Applies to the whole export. Whether a feature has a visible effect
        /// still depends on the font defining it.
        /// </summary>
        public GsubFeature GsubFeatures { get; set; } = GsubFeature.Liga | GsubFeature.Clig | GsubFeature.Calt | GsubFeature.Ccmp;

        /// <summary>
        /// GPOS (glyph positioning) features requested when shaping text, e.g. kerning and mark
        /// attachment. Applies to the whole export. Whether a feature has a visible effect still
        /// depends on the font defining it.
        /// </summary>
        public GposFeature GposFeatures { get; set; } = GposFeature.Kern | GposFeature.Mark;

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
    }
}