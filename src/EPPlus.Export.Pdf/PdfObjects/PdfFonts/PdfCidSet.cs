using EPPlus.Fonts.OpenType;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace EPPlus.Export.Pdf.PdfObjects.PdfFonts
{
    internal class PdfCidSet : PdfObject
    {
        byte[] CidSet;
        public PdfCidSet(int objectNumber, byte[] cidSet , int version = 0) : base(objectNumber, version)
        {
            CidSet = cidSet;
        }

        internal override string RenderDictionary()
        {
            return $"<< /Length {CidSet.Length} >>\n" + $"stream\n|BINARY DATA|\nendstream";
        }

        internal override void RenderDictionary(BinaryWriter bw)
        {
            WriteAscii(bw, $"<< /Length {CidSet.Length} >>\nstream\n");
            bw.Write(CidSet);
            WriteAscii(bw, "\nendstream");
        }
    }
}
