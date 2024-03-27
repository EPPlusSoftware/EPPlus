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
