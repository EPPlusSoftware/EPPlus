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
namespace EPPlusImageRenderer.RenderItems
{
    public enum PathCommandType : byte
    {
        Move = 0,
        Line = 1,
        HorizontalLine = 2,
        VerticalLine = 3,
        CubicBézier = 4,
        QuadraticBézier = 5,
        Arc = 6,
        End = 0xFF
    }

}