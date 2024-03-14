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
using OfficeOpenXml.Style;

namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IFont
    {
        internal string Name { get; }
        internal float Size { get; }
        internal IStyleColor Color { get; }
        internal bool HasValue { get; }
        internal bool Bold { get; }
        internal bool Italic { get; }
        internal bool Strike { get; }
        internal ExcelUnderLineType UnderLineType { get; }
    }
}
