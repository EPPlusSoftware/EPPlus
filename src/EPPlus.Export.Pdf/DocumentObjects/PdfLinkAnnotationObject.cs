using EPPlus.Export.Pdf.Helpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal class PdfLinkAnnotationObject : PdfObject
    {
        internal double LowerLeftX, LowerLeftY, UpperRightX, UpperRightY;
        internal string Uri;

        public PdfLinkAnnotationObject(int objectNumber, double lowerLeftX, double lowerLeftY, double upperRightX, double upperRightY, string uri, int version = 0)
            : base(objectNumber, version)
        {
            LowerLeftX = lowerLeftX;
            LowerLeftY = lowerLeftY;
            UpperRightX = upperRightX;
            UpperRightX = upperRightY;
            Uri = uri;
        }

        private string Body()
        {
            return $"<< /Type /Annot /Subtype /Link\n" +
                   $"   /Rect [ {LowerLeftX.ToPdfStringF4()} {LowerLeftY.ToPdfStringF4()} {UpperRightX.ToPdfStringF4()} {UpperRightY.ToPdfStringF4()} ]\n" +
                   $"   /Border [ 0 0 0 ]\n" +
                   $"   /F 4\n" +                                   // Print flag (PDF/A-friendly; Link needs no /AP)
                   $"   /A << /Type /Action /S /URI /URI {Uri.ToPdfLiteral()} >> >>";
        }

        internal override string RenderDictionary()
        {
            return Body();
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, Body());
        }
    }
}
