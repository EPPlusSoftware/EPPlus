using OfficeOpenXml.PDF.PdfGraphics;
using OfficeOpenXml.PDF.PdfObjects;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.PDF
{
    public class ExcelPdf
    {
        string header = "%PDF-1.7\n";
        List<PdfObject> body = new List<PdfObject>();
        List<PdfCrossRefTable> crossRefTable = new List<PdfCrossRefTable>();
        List<PdfTrailer> trailer = new List<PdfTrailer>();

        public readonly Dictionary<int, string> fontResources = new Dictionary<int, string>();

        public ExcelPdf()
        {
        }

        public void AddFont(string fontName = "Helvetica")
        {
            var font = new PdfFont(body.Count + 1, fontName);
            body.Add(font);
            fontResources.Add(body.IndexOf(font) + 1, "F" + (fontResources.Count + 1));
        }

        public void AddText(string text, string fontResourceName, int size, float x, float y)
        {
            var content = new PdfContentStream(body.Count + 1);
            content.AddText(fontResourceName, size, x, y, text);
            body.Add(content);
        }

        public void AddRectangle(float x, float y, float width, float height, PdfColor stroke = null, PdfColor fill = null)
        {
            var content = new PdfContentStream(body.Count + 1);
            content.AddRectangle(x, y, width, height, stroke != null ? true : false, fill != null ? true : false, stroke, fill);
            body.Add(content);
        }

        //create page
        private PdfPage AddPage(int pagesObjectNumber, List<int> contentObjectNumbers)
        {
            var page = new PdfPage(body.Count + 1, pagesObjectNumber, contentObjectNumbers, fontResources);
            body.Add(page);
            return page;
        }
        //create pages
        private PdfPages AddPages()
        {
            var pages = new PdfPages(body.Count + 1, new List<int>{});
            body.Add(pages);
            return pages;
        }
        //create Catalog
        private PdfCatalog AddCatalog(int pagesObjectNumber)
        {
            var catalog = new PdfCatalog(body.Count + 1, pagesObjectNumber);
            body.Add(catalog);
            return catalog;
        }

        public void CreatePdf(string Filename)
        {
            var pages = AddPages();
            List<int> contentObjectNumbers = new List<int>();
            contentObjectNumbers = body.OfType<PdfContentStream>().Select(con => con.objectNumber).ToList();
            var page = AddPage(pages.objectNumber, contentObjectNumbers);
            pages.pageObjectNumbers.Add(page.objectNumber);
            var catalog = AddCatalog(pages.objectNumber);
            using (var fs = new FileStream(Filename, FileMode.Create, FileAccess.Write))
            {
                var xrefPositions = new List<long>();

                // We'll use a BinaryWriter for precise control of bytes
                using (var writer = new BinaryWriter(fs, Encoding.ASCII))
                {
                    //Write header
                    writer.Write(Encoding.ASCII.GetBytes(header));
                    //Write body
                    foreach (var pdfobj in body)
                    {
                        xrefPositions.Add(fs.Position);
                        writer.Write(pdfobj.ToPdfBytes());

                    }
                    //Write CrossReference
                    long xrefStart = fs.Position;
                    writer.Write(Encoding.ASCII.GetBytes("xref\n"));
                    writer.Write(Encoding.ASCII.GetBytes($"0 {body.Count + 1}\n"));
                    writer.Write(Encoding.ASCII.GetBytes("0000000000 65535 f \n")); // Object 0 is always free
                    foreach (long pos in xrefPositions)
                    {
                        writer.Write(Encoding.ASCII.GetBytes(pos.ToString("D10") + " 00000 n \n"));
                    }
                    // Write trailer
                    writer.Write(Encoding.ASCII.GetBytes("trailer\n"));
                    writer.Write(Encoding.ASCII.GetBytes($"<< /Size {body.Count + 1} /Root {catalog.objectNumber} 0 R >>\n"));
                    writer.Write(Encoding.ASCII.GetBytes("startxref\n"));
                    writer.Write(Encoding.ASCII.GetBytes($"{xrefStart}\n"));
                    writer.Write(Encoding.ASCII.GetBytes("%%EOF\n"));
                }
            }
        }


        //public void WritePdf(string text)
        //{
        //    string outputPath = "HelloWorld.pdf";
        //    using (var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        //    {
        //        var xrefPositions = new List<long>();
        //        int objNumber = 1;

        //        // We'll use a BinaryWriter for precise control of bytes
        //        using (var writer = new BinaryWriter(fs, Encoding.ASCII))
        //        {
        //            // Write PDF Header
        //            writer.Write(Encoding.ASCII.GetBytes(header));

        //            // Object 1: Font
        //            xrefPositions.Add(fs.Position);
        //            WriteObject(writer, objNumber++, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

        //            // Object 2: Content stream (text drawing)
        //            string content = $"BT /F1 24 Tf 100 700 Td ({text}) Tj ET";
        //            byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        //            xrefPositions.Add(fs.Position);
        //            writer.Write(Encoding.ASCII.GetBytes($"{objNumber} 0 obj\n"));
        //            writer.Write(Encoding.ASCII.GetBytes($"<< /Length {contentBytes.Length} >>\nstream\n"));
        //            writer.Write(contentBytes);
        //            writer.Write(Encoding.ASCII.GetBytes("\nendstream\nendobj\n"));
        //            objNumber++;

        //            // Object 3: Page
        //            xrefPositions.Add(fs.Position);
        //            WriteObject(writer, objNumber++, "<< /Type /Page /Parent 4 0 R /Resources << /Font << /F1 1 0 R >> >> /Contents 2 0 R /MediaBox [0 0 595 842] >>");

        //            // Object 4: Pages
        //            xrefPositions.Add(fs.Position);
        //            WriteObject(writer, objNumber++, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");

        //            // Object 5: Catalog
        //            xrefPositions.Add(fs.Position);
        //            WriteObject(writer, objNumber++, "<< /Type /Catalog /Pages 4 0 R >>");

        //            // Save xref offset
        //            long xrefStart = fs.Position;

        //            // Write XREF table
        //            writer.Write(Encoding.ASCII.GetBytes("xref\n"));
        //            writer.Write(Encoding.ASCII.GetBytes($"0 {objNumber}\n"));
        //            writer.Write(Encoding.ASCII.GetBytes("0000000000 65535 f \n")); // Object 0 is always free
        //            foreach (long pos in xrefPositions)
        //            {
        //                writer.Write(Encoding.ASCII.GetBytes(pos.ToString("D10") + " 00000 n \n"));
        //            }

        //            // Write trailer
        //            writer.Write(Encoding.ASCII.GetBytes("trailer\n"));
        //            writer.Write(Encoding.ASCII.GetBytes($"<< /Size {objNumber} /Root 5 0 R >>\n"));
        //            writer.Write(Encoding.ASCII.GetBytes("startxref\n"));
        //            writer.Write(Encoding.ASCII.GetBytes($"{xrefStart}\n"));
        //            writer.Write(Encoding.ASCII.GetBytes("%%EOF\n"));
        //        }
        //    }

        //    static void WriteObject(BinaryWriter writer, int objNumber, string content)
        //    {
        //        writer.Write(Encoding.ASCII.GetBytes($"{objNumber} 0 obj\n"));
        //        writer.Write(Encoding.ASCII.GetBytes(content));
        //        writer.Write(Encoding.ASCII.GetBytes("\nendobj\n"));
        //    }
    }
}
