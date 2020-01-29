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
    internal class FlagCellStore : CellStore<byte>
    {
        internal void SetFlagValue(int Row, int Col, bool value, CellFlags cellFlags)
        {
            CellFlags currentValue = (CellFlags)GetValue(Row, Col);
            if (value)
            {
                SetValue(Row, Col, (byte)(currentValue | cellFlags)); // add the CellFlag bit
            }
            else
            {
                SetValue(Row, Col, (byte)(currentValue & ~cellFlags)); // remove the CellFlag bit
            }
        }
        internal bool GetFlagValue(int Row, int Col, CellFlags cellFlags)
        {
            return !(((byte)cellFlags & GetValue(Row, Col)) == 0);
        }
    }
}