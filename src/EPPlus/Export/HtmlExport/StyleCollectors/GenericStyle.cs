using OfficeOpenXml.Style.Dxf;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class GenericStyle
    {
        internal GenericFill Fill { get; set; } = null;

        internal GenericFont Font { get; set; } = null;

        internal GenericBorder Border { get; set; } = null;


        internal GenericStyle() 
        {
            
        }

        internal GenericStyle(ExcelDxfStyleConditionalFormatting style)
        {
            Fill = new GenericFill(style.Fill);
        }

        internal GenericStyle(ExcelXfs style)
        {
            if(style.FillId > 0)
            {
                Fill = new GenericFill(style.Fill);
            }
        }

    }
}
