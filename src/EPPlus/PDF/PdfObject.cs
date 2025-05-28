using System.Text;

namespace OfficeOpenXml.PDF
{
    internal abstract class PdfObject
    {
        internal int objectNumber;
        internal int version;
        internal long byteOffset;

        internal virtual bool HasStream => false;
        internal virtual byte[] StreamData => null;

        public PdfObject(int objectNumber, int version = 0)
        {
            this.objectNumber = objectNumber;
            this.version = version;
        }

        public virtual byte[] ToPdfBytes()
        {
            var sb = new StringBuilder();
            sb.AppendFormat("{0} {1} obj\n", objectNumber, version);
            sb.Append(RenderDictionary());
            sb.Append("\nendobj\n");
            return Encoding.ASCII.GetBytes(sb.ToString());
        }

        internal abstract string RenderDictionary();
    }
}
