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
using System;

namespace OfficeOpenXml.Drawing
{
    /// <summary>
    /// Modifiers for a style entry
    /// </summary>
    [Flags]
    public enum eStyleEntryModifier
    {
        /// <summary>
        /// This style entry can be replaced with no fill
        /// </summary>
        AllowNoFillOverride = 1,
        /// <summary>
        /// This style entry can be replaced with no line
        /// </summary>
        AllowNoLineOverride = 2,
    }
}
