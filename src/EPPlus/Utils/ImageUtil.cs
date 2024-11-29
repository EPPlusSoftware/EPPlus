using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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
    }
}
