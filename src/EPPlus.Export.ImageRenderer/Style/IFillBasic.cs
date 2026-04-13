using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Theme;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Style
{
    public interface IFillBasic
    {
        /// <summary>
        /// Fill type of object
        /// </summary>
        eFillStyle Style { get; }

        string GetBackgroundColor(ExcelTheme theme);
        string GetGradientColor1(ExcelTheme theme);
        string GetGradientColor2(ExcelTheme theme);
    }
}
