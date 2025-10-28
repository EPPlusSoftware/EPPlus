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
using System;

namespace EPPlusImageRenderer.RenderItems
{
    internal static class SvgExtensions
    {
        internal static char AsCommandChar(this PathCommandType type)
        {
            switch (type)
            {
                case PathCommandType.Move:
                    return 'M';
                case PathCommandType.Line:
                    return 'L';
                case PathCommandType.HorizontalLine:
                    return 'H';
                case PathCommandType.VerticalLine:
                    return 'V';
                case PathCommandType.CubicBézier:
                    return 'C';
                case PathCommandType.QuadraticBézier:
                    return 'Q';
                case PathCommandType.Arc:
                    return 'A';
                case PathCommandType.End:
                    return 'Z';                
                default:
                    throw new NotImplementedException("SVG path type not implemented");
            }
        }
    }

}