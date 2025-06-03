namespace FontLab1.Tables.Post
{
    internal class PostTable
    {
        public int version { get; set; }
        public double italicAngle { get; set; }
        public short underlinePosition {  get; set; }
        public short underlineThickness { get; set; }
        public uint isFixedPitch { get; set; }
        public uint minMemType42 { get; set; }
        public uint maxMemType42 { get; set; }
        public uint minMemType1 { get; set; }
        public uint maxMemType1 { get;set; }
    }
}
