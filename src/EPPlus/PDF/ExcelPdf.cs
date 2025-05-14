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

        public ExcelPdf()
        {

        }

        public void WritePdf(string text)
        {
            string outputPath = "HelloWorld.pdf";
            using (var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
            {
                var xrefPositions = new List<long>();
                int objNumber = 1;

                // We'll use a BinaryWriter for precise control of bytes
                using (var writer = new BinaryWriter(fs, Encoding.ASCII))
                {
                    // Write PDF Header
                    writer.Write(Encoding.ASCII.GetBytes(header));

                    // Object 1: Font
                    xrefPositions.Add(fs.Position);
                    WriteObject(writer, objNumber++, "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>");

                    // Object 2: Content stream (text drawing)
                    string content = $"BT /F1 24 Tf 100 700 Td ({text}) Tj ET";
                    byte[] contentBytes = Encoding.ASCII.GetBytes(content);
                    xrefPositions.Add(fs.Position);
                    writer.Write(Encoding.ASCII.GetBytes($"{objNumber} 0 obj\n"));
                    writer.Write(Encoding.ASCII.GetBytes($"<< /Length {contentBytes.Length} >>\nstream\n"));
                    writer.Write(contentBytes);
                    writer.Write(Encoding.ASCII.GetBytes("\nendstream\nendobj\n"));
                    objNumber++;

                    // Object 3: Page
                    xrefPositions.Add(fs.Position);
                    WriteObject(writer, objNumber++, "<< /Type /Page /Parent 4 0 R /Resources << /Font << /F1 1 0 R >> >> /Contents 2 0 R /MediaBox [0 0 595 842] >>");

                    // Object 4: Pages
                    xrefPositions.Add(fs.Position);
                    WriteObject(writer, objNumber++, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>");

                    // Object 5: Catalog
                    xrefPositions.Add(fs.Position);
                    WriteObject(writer, objNumber++, "<< /Type /Catalog /Pages 4 0 R >>");

                    // Save xref offset
                    long xrefStart = fs.Position;

                    // Write XREF table
                    writer.Write(Encoding.ASCII.GetBytes("xref\n"));
                    writer.Write(Encoding.ASCII.GetBytes($"0 {objNumber}\n"));
                    writer.Write(Encoding.ASCII.GetBytes("0000000000 65535 f \n")); // Object 0 is always free
                    foreach (long pos in xrefPositions)
                    {
                        writer.Write(Encoding.ASCII.GetBytes(pos.ToString("D10") + " 00000 n \n"));
                    }

                    // Write trailer
                    writer.Write(Encoding.ASCII.GetBytes("trailer\n"));
                    writer.Write(Encoding.ASCII.GetBytes($"<< /Size {objNumber} /Root 5 0 R >>\n"));
                    writer.Write(Encoding.ASCII.GetBytes("startxref\n"));
                    writer.Write(Encoding.ASCII.GetBytes($"{xrefStart}\n"));
                    writer.Write(Encoding.ASCII.GetBytes("%%EOF\n"));
                }
            }

            static void WriteObject(BinaryWriter writer, int objNumber, string content)
            {
                writer.Write(Encoding.ASCII.GetBytes($"{objNumber} 0 obj\n"));
                writer.Write(Encoding.ASCII.GetBytes(content));
                writer.Write(Encoding.ASCII.GetBytes("\nendobj\n"));
            }
        }
    }
}
