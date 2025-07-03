namespace OfficeOpenXml.PDF.Pdfhelpers
{
    internal static class PdfUnits
    {
        public const double PointsPerInch = 72.0;
        public const double MmPerInch = 25.4;

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
            double pixels = System.Math.Truncate(7 * columnWidth + 5);
            double points = pixels * 0.75;
            return points;
        }

    }
}
