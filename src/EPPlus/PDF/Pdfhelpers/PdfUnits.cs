namespace OfficeOpenXml.PDF.Pdfhelpers
{
    internal static class PdfUnits
    {
        public const double PointsPerInch = 72.0d;
        public const double MmPerInch = 25.4d;
        public const double DPI = 600d;

        public static double MmToPoints(double mm)
        {
            return mm * PointsPerInch / MmPerInch;
        }

        public static double PointsToMm(double points)
        {
            return points * MmPerInch / PointsPerInch;
        }

        public static int MmToPointsRounded(double mm)
        {
            return (int)System.Math.Round(MmToPoints(mm));
        }

        public static int PointsToMmRounded(double points)
        {
            return (int)System.Math.Round(PointsToMm(points));
        }

        public static double ExcelColumnWidthToPoints(double columnWidth)
        {
            double pixels = System.Math.Truncate(7 * columnWidth + 6); //These values are guessed.
            double points = pixels * 0.75d; //These values are guessed.
            return points;
        }

        public static double ExcelRowHeightToPoints(double rowHeight)
        {
            return rowHeight + 0.25d; //These values are guessed.
        }

    }
}
