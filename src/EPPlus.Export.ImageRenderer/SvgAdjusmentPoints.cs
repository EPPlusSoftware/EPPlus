/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  27/11/2025         EPPlus Software AB           EPPlus 9
 *************************************************************************************************/

using System.Collections.Generic;

namespace EPPlusImageRenderer
{
    internal class SvgAdjustmentPoint
    {
        public int ItemIndex { get; set; }
        public List<SvgCommand> Commands { get; set; } = null;
        public AdjustmentType AdjustmentType { get; set; }
    }
}
