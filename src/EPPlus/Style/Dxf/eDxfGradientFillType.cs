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
namespace OfficeOpenXml
{
    /// <summary>
    /// Type of gradient fill
    /// </summary>
    public enum eDxfGradientFillType
    {
        /// <summary>
        /// Linear gradient type. Linear gradient type means that the transition from one color to the next is along a line.
        /// </summary>
        Linear,
        /// <summary>
        /// Path gradient type. Path gradient type means the that the transition from one color to the next is a rectangle, defined by coordinates.
        /// </summary>
        Path
    }
}