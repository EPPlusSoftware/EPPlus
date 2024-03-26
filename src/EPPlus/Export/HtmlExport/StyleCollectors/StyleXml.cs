/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  03/14/2024         EPPlus Software AB           Epplus 7.1
 *************************************************************************************************/
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
            _style = style;

            if (style.FillId >=  0)
            {
                Fill = new FillXml(style.Fill);
            }
            if(style.FontId >= 0)
            {
                Font = new FontXml(style.Font);
            }
            if(style.BorderId >= 0) 
            {
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
