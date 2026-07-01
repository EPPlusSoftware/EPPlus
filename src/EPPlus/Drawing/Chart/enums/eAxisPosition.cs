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
namespace OfficeOpenXml.Drawing.Chart
{
    /// <summary>
    /// Position of the axis.
    /// </summary>
    public enum eAxisPosition
    {
        /// <summary>
        /// Left
        /// </summary>
        Left = 0,
        /// <summary>
        /// Bottom
        /// </summary>
        Bottom = 1,
        /// <summary>
        /// Right
        /// </summary>
        Right = 2,
        /// <summary>
        /// Top
        /// </summary>
        Top = 3
    }
    public enum eActualAxisPosition
    {
        /// <summary>
        /// Left
        /// </summary>
        Left = 0,
        /// <summary>
        /// Bottom
        /// </summary>
        Bottom = 1,
        /// <summary>
        /// Right
        /// </summary>
        Right = 2,
        /// <summary>
        /// Top
        /// </summary>
        Top = 3,
        /// <summary>
        /// If there are two axis on the left side, this is the second axis left (most to the left)
        /// </summary>
        LeftSecond = 5,
        /// <summary>
        /// If there are two axis on the right side, this is the second axis right (most to the right)
        /// </summary>
        RightSecond = 7,
        /// <summary>
        /// If there are two axis on the top, this is the second top left (most to the top)
        /// </summary>
        TopSecond = 9,
        /// <summary>
        /// If there are two axis on the bottom, this is the second bottom left (most to the top)
        /// </summary>
        BottomSecond = 11
    }
}