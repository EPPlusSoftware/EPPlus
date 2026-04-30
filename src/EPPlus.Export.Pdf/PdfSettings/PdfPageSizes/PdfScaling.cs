namespace EPPlus.Export.Pdf.PdfSettings.PdfPageSizes
{
    public enum ScalingMode
    {
        /// <summary>
        /// Adjust size based on the normal size
        /// </summary>
        AdjustToNormalSize,
        /// <summary>
        /// Adjust size based on number of pages
        /// </summary>
        FitToPages
    }

    public class PdfScaling
    {
        public ScalingMode ScalingMode { get; set; }

        /// <summary>
        /// Used when <see cref="ScalingMode"/> is set to <c>AdjustToNormalSize</c>.
        /// The scale factor is expressed as a multiplier, where 1.0 represents 100% (no scaling), 0.5 represents 50%, and 2.0 represents 200%.
        /// </summary>
        public double Scale { get; set; } = 1.0;

        /// <summary>
        /// Used when <see cref="ScalingMode"/> is set to <c>FitToPages</c>.
        /// Scales the content to fit to the specifed number of pages in width.
        /// </summary>
        public int PagesWide { get; set; } = 1;
        /// <summary>
        /// Used when <see cref="ScalingMode"/> is set to <c>FitToPages</c>.
        /// Scales the content to fit to the specifed number of pages in height.
        /// </summary>
        public int PagesTall { get; set; } = 1;

        public PdfScaling(double scale)
        {
            ScalingMode = ScalingMode.AdjustToNormalSize;
            Scale = scale;
        }

        public PdfScaling(int pagesWide, int pagesTall)
        {
            ScalingMode = ScalingMode.FitToPages;
            PagesWide = pagesWide;
            PagesTall = pagesTall;
        }


        public static PdfScaling NoScaling => new PdfScaling(1d);
        public static PdfScaling FitSheetToOnePage => new PdfScaling(1, 1);
        public static PdfScaling FitAllColumnsOnOnePage => new PdfScaling(1, 0);
        public static PdfScaling FitAllRowsOnOnePage => new PdfScaling(0, 1);

    }

}
