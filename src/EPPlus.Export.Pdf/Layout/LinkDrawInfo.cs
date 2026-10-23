namespace EPPlus.Export.Pdf.Layout
{
    internal struct LinkDrawInfo
    {
        public double X, Y, Width, Height;
        public string Uri;
        public string InternalAddress;
        public string TargetSheet;
        public int TargetRow;
        public int TargetColumn;
    }
}
