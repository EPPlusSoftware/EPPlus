namespace OfficeOpenXml.PDF.PdfResources
{
    internal class PdfResource
    {
        internal readonly string labelPrefix;
        internal int labelNumber;

        internal string Label
        {
            get
            {
                return labelPrefix + labelNumber;
            }
        }

        public PdfResource(string labelPrefix, int labelNumber)
        {
            this.labelPrefix = labelPrefix;
            this.labelNumber = labelNumber;
        }
    }
}
