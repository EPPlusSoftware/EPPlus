/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/
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