using OfficeOpenXml.Drawing;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class BorderDrawing : IDrawingBorder
    {
        ExcelDrawingBorder _border;

        internal BorderDrawing(ExcelDrawingBorder border)
        {
            _border = border;
            Stroke = new FillDrawingBasic(border.Fill);
        }

        public IFill Stroke { get; private set; }
    }
}
