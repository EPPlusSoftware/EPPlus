/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  16/11/2026         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
using EPPlus.Export.Pdf;
using EPPlus.Export.Pdf.Settings;
using OfficeOpenXml.Utils.FileUtils;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeOpenXml.Export.PdfExport
{
    internal partial class EPPlusToPdfWriter
    {
        /// <summary>
        /// Builds the PDF synchronously and writes it to the given file asynchronously.
        /// </summary>
        public async Task SaveAsync(string fileName, CancellationToken cancellationToken)
        {
            // Validate before building, so an invalid path is reported before the (possibly long) layout build.
            var fullPath = FileHelper.ValidateOutputFilePath(fileName, nameof(fileName));
            BeginSave();
            try
            {
                var layout = _build();
                await new ExcelPdf().CreatePdfAsync(PdfDocumentSettings.From(_pageSettings), _dictionaries, layout, fullPath, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                DeleteTemporarySheets();
            }
        }

        /// <summary>
        /// Builds the PDF synchronously and writes it to the given stream asynchronously.
        /// </summary>
        public async Task SaveAsync(Stream stream, CancellationToken cancellationToken)
        {
            BeginSave();
            try
            {
                var layout = _build();
                await new ExcelPdf().CreatePdfAsync(PdfDocumentSettings.From(_pageSettings), _dictionaries, layout, stream, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                DeleteTemporarySheets();
            }
        }
    }
}