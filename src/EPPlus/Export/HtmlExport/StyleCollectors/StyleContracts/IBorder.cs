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
namespace OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts
{
    public interface IBorder
    {
        internal bool HasValue { get; }
        internal IBorderItem Top { get; }
        internal IBorderItem Bottom { get; }
        internal IBorderItem Left { get; }
        internal IBorderItem Right { get; }
    }
}
