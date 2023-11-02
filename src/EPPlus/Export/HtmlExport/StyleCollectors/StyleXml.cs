using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using OfficeOpenXml.Style;
using OfficeOpenXml.Style.XmlAccess;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors
{
    internal class StyleXml : IStyleExport
    {
        internal ExcelXfs _style;

        public string StyleKey 
        { 
            get
            {
                return GetStyleKey();
            }
        }

        public bool HasStyle
        {
            get
            {
                return _style.FontId > 0 ||
                   _style.FillId > 0 ||
                   _style.BorderId > 0 ||
                   _style.HorizontalAlignment != ExcelHorizontalAlignment.General ||
                   _style.VerticalAlignment != ExcelVerticalAlignment.Bottom ||
                   _style.TextRotation != 0 ||
                   _style.Indent > 0 ||
                   _style.WrapText;
            }
        }

        public IFill Fill { get; } = null;

        public IBorder Border { get; } = null;

        public IFont Font { get; } = null;

        public StyleXml(ExcelXfs style)        
        {
            if(style.FillId >  0)
            {
                _style = style;

                Fill = new FillXml(style.Fill);
                Font = new FontXml(style.Font);
                Border = new BorderXml(style.Border);
            }
        }

        internal string GetStyleKey()
        {
            var fbfKey = ((ulong)(uint)_style.FontId << 32 | (uint)_style.BorderId << 16 | (uint)_style.FillId);
            return fbfKey.ToString() + "|" + ((int)_style.HorizontalAlignment).ToString() + "|" + ((int)_style.VerticalAlignment).ToString() + "|" + _style.Indent.ToString() + "|" + _style.TextRotation.ToString() + "|" + (_style.WrapText ? "1" : "0");
        }
    }
}
