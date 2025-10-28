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
using System.Linq;

namespace EPPlusImageRenderer
{
    internal struct SvgCommand
    {
        public SvgCommand(int itemIndex)
        {
            Index = itemIndex;
            Coordinates = new Dictionary<short, SvgCoordinate>();
        }
        public SvgCommand(int itemIndex, params SvgCoordinate[] coordinates)
        {
            Index = itemIndex;
            Coordinates = coordinates.ToDictionary(x=>x.Value, y=>y);
        }
        public int Index { get; set; }
        public Dictionary<short, SvgCoordinate> Coordinates { get; set; }
    }
}