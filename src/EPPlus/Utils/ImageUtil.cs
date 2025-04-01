using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml.Drawing;

namespace OfficeOpenXml.Utils
{
    internal static class ImageUtil
    {
        internal static void ResizeImageWithMaxSize(double maxWidth, double maxHeight, int width, int height, out int newWidth, out int newHeight)
        {
            double xRatio = maxWidth / (double)width;
            double yRatio = maxHeight / (double)height;

            double ratio = xRatio < yRatio ? xRatio : yRatio;

            newWidth = Convert.ToInt32(width * ratio);
            newHeight = Convert.ToInt32(height * ratio);
        }
        internal static void CalculateDPI(ExcelImage Image, float Standard_DPI, ref double horizontalResoluton, ref double verticalResolution)
        {
            var horizontalDpi = Image.Bounds.HorizontalResolution == 0 ? Standard_DPI : Image.Bounds.HorizontalResolution;
            var verticalDpi = Image.Bounds.VerticalResolution == 0 ? Standard_DPI : Image.Bounds.VerticalResolution;
            horizontalResoluton = Image.Bounds.Width / (horizontalDpi / Standard_DPI);
            verticalResolution = Image.Bounds.Height / (verticalDpi / Standard_DPI);
        }
        /// <summary>
        /// Converts from pixels to points at a specified DPI (usually 72)
        /// </summary>
        /// <param name="pixelValue"></param>
        /// <param name="standard_DPI"></param>
        /// <param name="resolution">Horizontal or Vertical Resolution</param>
        /// <returns></returns>
        internal static double PixelToPointConversion(double pixelValue, double resolution, float standard_DPI = 72) //Pixel --> Points
        {
            double pointsValue = pixelValue * standard_DPI / resolution;
            return pointsValue;
        }
    }
}
