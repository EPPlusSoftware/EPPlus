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
namespace OfficeOpenXml.Core.CellStore
{
    /// <summary>
    /// For cell value structure (for memory optimization of huge sheet)
    /// </summary>
    internal struct ExcelValue
    {
        internal object _value;
        internal int _styleId;

        public override string ToString()
        {
            if (_value != null) return _value.ToString();
            return "null";
        }
    }
}