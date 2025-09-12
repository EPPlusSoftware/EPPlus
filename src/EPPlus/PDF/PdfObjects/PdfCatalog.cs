namespace OfficeOpenXml.PDF.PdfObjects
{
    internal class PdfCatalog : PdfObject
    {
        private readonly int pagesObjectNumber;

        public PdfCatalog(int objectNumber, int pagesObjectNumber, int version = 0)
            : base(objectNumber, version)
        {
            this.pagesObjectNumber = pagesObjectNumber;
        }

        internal override string RenderDictionary()
        {
            return $"<< /Type /Catalog\n" +
                   $"   /Pages {pagesObjectNumber} 0 R >>";
        }
    }
}
