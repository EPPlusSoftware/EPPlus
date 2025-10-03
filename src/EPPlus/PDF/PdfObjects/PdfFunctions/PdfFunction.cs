namespace OfficeOpenXml.PDF.PdfObjects.PdfFunctions
{
    internal abstract class PdfFunction : PdfObject
    {
        //    Implemented    Function type
        //    [ ]            0 Sampled function
        //    [X]            2 Exponential interpolation function
        //    [X]            3 Stitching function
        //    [ ]            4 PostScript calculator function

        internal double[] Domain;
        internal double[] Range;

        public PdfFunction(int objectNumber, int version = 0) : base(objectNumber, version) { }
    }
}
