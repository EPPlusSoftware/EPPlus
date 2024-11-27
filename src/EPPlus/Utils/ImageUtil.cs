using OfficeOpenXml.Drawing.Style.Coloring;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Utils
{
    internal static class ImageUtil
    {
       internal static void ResizeImageWithMaxSize(int maxWidth, int maxHeight, int width, int height, out int newHeight, out int newWidth)
        {
            double xRatio = (double)maxWidth / (double)width;
            double yRatio = (double)maxHeight / (double)height;

            double ratio = xRatio < yRatio ? xRatio : yRatio;

            newHeight = Convert.ToInt32(width * ratio);
            newWidth = Convert.ToInt32(height * ratio);
        }
    }
}
