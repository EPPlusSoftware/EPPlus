using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class StyleDxf : IStyleExport
    {
        ExcelDxfStyleConditionalFormatting _style;

        public bool HasStyle
        {
            get { return _style.HasValue; }
        }

        public string StyleKey { get { return _style.Id; } }

        public IFill Fill { get; } = null;
        public IFont Font { get; } = null;
        public IBorder Border { get; } = null;

        public StyleDxf(ExcelDxfStyleConditionalFormatting style)
        {
            _style = style;

            if (style.Fill != null)
            {
                Fill = new FillDxf(style.Fill);
            }
            if (style.Font != null)
            {
                Font = new FontDxf(style.Font);
            }
            if (style.Border != null)
            {
                Border = new BorderDxf(style.Border);
            }
        }

        public StyleDxf(ExcelDxfStyleLimitedFont style)
        {
            _style = style.ToDxfConditionalFormattingStyle();

            if (style.Fill != null)
            {
                Fill = new FillDxf(style.Fill);
            }
            if (style.Font != null)
            {
                Font = new FontDxf(style.Font);
            }
            if (style.Border != null)
            {
                Border = new BorderDxf(style.Border);
            }
        }
    }
}
