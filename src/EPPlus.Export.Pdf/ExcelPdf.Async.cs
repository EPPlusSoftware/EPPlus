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
using EPPlus.Export.Pdf.DocumentObjects;
using EPPlus.Export.Pdf.Resources;
using EPPlus.Export.Pdf.Settings;
using EPPlus.Graphics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf
{
    internal partial class ExcelPdf
    {
        //Creates the Pdf document asynchronously. The document is built synchronously; only the writing is asynchronous.
        internal async Task CreatePdfAsync(PdfDocumentSettings documentSettings, PdfDictionaries dictionaries, Transform layout, string fileName, CancellationToken cancellationToken)
        {
            // Check before the file is created, so an already cancelled call does not touch the file.
            cancellationToken.ThrowIfCancellationRequested();
            var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write, FileShare.None, 4096, true);
            try
            {
                using (fs)
                {
                    await CreatePdfAsync(documentSettings, dictionaries, layout, fs, cancellationToken).ConfigureAwait(false);
                }
            }
            catch
            {
                // Don't leave a truncated PDF behind, also on cancellation.
                TryDeleteFile(fileName);
                throw;
            }
            // Debug output only, written synchronously.
            WriteDebugText(fileName);
        }

        //Creates the Pdf document asynchronously. The document is built synchronously; only the writing is asynchronous.
        internal async Task CreatePdfAsync(PdfDocumentSettings documentSettings, PdfDictionaries dictionaries, Transform layout, Stream stream, CancellationToken cancellationToken)
        {
            ValidateStream(stream);
            cancellationToken.ThrowIfCancellationRequested();
            PdfCatalog catalog;
            PdfInfoObject info;
            BuildDocument(documentSettings, dictionaries, layout, out catalog, out info);

            var crossRefTable = new PdfCrossRefTable();
            long offset = 0;
            using (var buffer = new MemoryStream())
            using (var bw = new BinaryWriter(buffer, Encoding.ASCII))
            {
                var writer = new PdfWriter(bw, _debugText != null, documentSettings.CompressionLevel);
                WriteHeader(writer);
                offset += await FlushBufferAsync(bw, buffer, stream, cancellationToken).ConfigureAwait(false);

                foreach (var pdfObj in _document)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    crossRefTable.AddPosition(offset);
                    WriteObject(writer, pdfObj);
                    offset += await FlushBufferAsync(bw, buffer, stream, cancellationToken).ConfigureAwait(false);
                }

                WriteCrossRefAndTrailer(writer, crossRefTable, offset, catalog, info);
                await FlushBufferAsync(bw, buffer, stream, cancellationToken).ConfigureAwait(false);
                _debugText?.Append(writer.DebugText);
            }
            await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
        }

        // Copies the buffered bytes to the target, resets the buffer and returns the number of bytes written.
        private static async Task<long> FlushBufferAsync(BinaryWriter bw, MemoryStream buffer, Stream target, CancellationToken cancellationToken)
        {
            bw.Flush();
            long length = buffer.Length;
            await target.WriteAsync(buffer.GetBuffer(), 0, (int)length, cancellationToken).ConfigureAwait(false);
            buffer.Position = 0;
            buffer.SetLength(0);
            return length;
        }
    }
}