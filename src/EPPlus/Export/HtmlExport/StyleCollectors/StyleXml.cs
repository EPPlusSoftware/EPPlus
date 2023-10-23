using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class StyleXml : IStyle
    {
        public IFill Fill { get; } = null;

        public IBorder Border { get; } = null;

        public IFont Font { get; } = null;

        public StyleXml(ExcelXfs style)        
        {
            if(style.FillId >  0)
            {
                Fill = new FillXml(style.Fill);
                Font = new FontXml(style.Font);
            }
        }
    }
}
