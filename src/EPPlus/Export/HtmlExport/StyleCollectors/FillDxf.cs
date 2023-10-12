using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class FillDxf : IFill
    {
        ExcelDxfFill _fill;

        public FillDxf(ExcelDxfFill fill)
        {
            _fill = fill;
        }

        public ExcelFillStyle PatternType 
        { 
            get 
            {
                if (_fill.HasValue)
                {
                    return _fill.PatternType.Value;
                }

                return ExcelFillStyle.None;
            } 
        }
    }
}
