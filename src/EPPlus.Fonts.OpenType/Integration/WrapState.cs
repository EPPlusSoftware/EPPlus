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
    internal class WrapState
    {
        public WrapState(double lineWidth, double spaceWidth)
        {
            CurrentLineWidth = lineWidth;
            SpaceWidth = spaceWidth;
        }

        public int LineStart { get; set; }
        public int WordStart { get; set; }
        public double CurrentLineWidth { get; set; }
        public double CurrentWordWidth { get; set; }
        public double SpaceWidth { get; set; }

        public bool IsCompleteWordReady(CharacterType charType, int currentPosition)
        {
            return (charType == CharacterType.Space || charType == CharacterType.EndOfText)
                   && WordStart < currentPosition; 
        }
    }
}