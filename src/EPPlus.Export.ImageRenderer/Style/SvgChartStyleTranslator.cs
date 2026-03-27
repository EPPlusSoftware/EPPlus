using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EPPlus.Export.ImageRenderer.Style
{
    internal class SvgChartStyleTranslator : IStyleExportDrawing
    {
        public SvgChartStyleTranslator(ExcelChartStyle chartStyle) 
        { 

        }
        public string StyleKey => throw new NotImplementedException();

        public bool HasStyle => throw new NotImplementedException();

        public IFillBasic Fill => throw new NotImplementedException();
    }
}
