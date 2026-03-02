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
    internal enum RenderItemType
    {
        Path = 0,
        Rect = 1,
        Group = 2,
        Line = 3,
        Ellipse = 4,
        Text = 5,
        TSpan = 6,
        Paragraph = 7,
        CommentTitle = 8,
    }
}