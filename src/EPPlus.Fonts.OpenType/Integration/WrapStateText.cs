/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/20/2025         EPPlus Software AB           TextLayoutEngine implementation
 *************************************************************************************************/
using System.Collections.Generic;

namespace EPPlus.Fonts.OpenType.Integration
{
    internal class WrapStateText : WrapStateBase
    {
        public WrapStateText(double lineWidth, double spaceWidth)
        {
            CurrentLineWidth = lineWidth;
            SpaceWidth = spaceWidth;
        }

        public double SpaceWidth { get; set; }
    }
}