using EPPlus.Export.Pdf.Helpers;
using System.IO;

namespace EPPlus.Export.Pdf.DocumentObjects
{
    internal abstract class PdfStreamObject : PdfObject
    {
        protected PdfStreamObject(int objectNumber, int version = 0) : base(objectNumber, version) { }

        protected abstract (string dictionary, byte[] bytes) BuildStream();

        internal sealed override void RenderDictionary(BinaryWriter bw)
        {
            var (dictionary, bytes) = BuildStream();
            WriteAscii(bw, dictionary + "\nstream\n");
            bw.Write(bytes);
            WriteAscii(bw, "\nendstream");
        }

        internal override string RenderDictionary()
        {
            var (dictionary, bytes) = BuildStream();
            return dictionary + $"\nstream\n|{bytes.Length.ToPdfStringF0()} bytes binary data|\nendstream";
        }
    }
}
