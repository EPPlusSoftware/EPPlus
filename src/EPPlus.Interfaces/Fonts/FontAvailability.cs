/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/07/2025         EPPlus Software AB           EPPlus.Fonts.OpenType 1.0
 *************************************************************************************************/

namespace OfficeOpenXml.Interfaces.Fonts
{
    /// <summary>
    /// Result of a font availability check.
    /// </summary>
    public enum FontAvailability
    {
        /// <summary>The font family is not available on the system.</summary>
        NotFound,

        /// <summary>The font family is available, but not in the requested subfamily 
        /// (e.g. Regular exists but Bold was requested).</summary>
        FamilyOnly,

        /// <summary>The exact font family and subfamily is available.</summary>
        Exact
    }
}