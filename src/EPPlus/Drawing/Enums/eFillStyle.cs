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
namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// The Fillstyle
    /// </summary>
    public enum eFillStyle
    {
        /// <summary>
        /// No fill
        /// </summary>
        NoFill,
        /// <summary>
        /// A solid fill
        /// </summary>
        SolidFill,
        /// <summary>
        /// A smooth gradual transition from one color to the next
        /// </summary>
        GradientFill,
        /// <summary>
        /// A preset pattern  fill
        /// </summary>
        PatternFill,
        /// <summary>
        /// Picturefill
        /// </summary>
        BlipFill,
        /// <summary>
        /// Inherited fill from the parent in the group.
        /// </summary>
        GroupFill
    }
}