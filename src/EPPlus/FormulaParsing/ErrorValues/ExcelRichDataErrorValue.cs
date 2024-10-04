/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB       Initial release EPPlus 5
 *************************************************************************************************/
namespace OfficeOpenXml
{
    /// <summary>
    /// Represents spill errors 
    /// </summary>
    public class ExcelRichDataErrorValue : ExcelErrorValue
    {
        internal ExcelRichDataErrorValue(int rowOffset, int colOffset) : base(eErrorType.Spill)
        {
            SpillRowOffset = rowOffset;
            SpillColOffset = colOffset;
        }
        internal int SpillRowOffset { get; set; }
        internal int SpillColOffset { get; set; }
        internal bool IsPropagated
        {
            get;
            set;
        }
    }
}
