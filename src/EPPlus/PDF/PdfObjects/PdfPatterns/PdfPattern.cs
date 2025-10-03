namespace OfficeOpenXml.PDF.PdfObjects.PdfPatterns
{
    internal abstract class PdfPattern : PdfObject
    {
        public PdfPattern(int objectNumber, int version = 0) : base(objectNumber, version) { }
    }
}
