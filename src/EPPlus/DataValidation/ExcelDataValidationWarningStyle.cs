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

namespace OfficeOpenXml.DataValidation
{
    /// <summary>
    /// warning style, controls how Excel will handle invalid changes.
    /// </summary>
    public enum ExcelDataValidationWarningStyle
    {
        /// <summary>
        /// warning style will be excluded.
        /// Excel will default this to Stop warning style.
        /// </summary>
        undefined,
        /// <summary>
        /// stop warning style, invalid changes will not be accepted
        /// </summary>
        stop,
        /// <summary>
        /// warning will be presented when an attempt to an invalid change is done, but the change will be accepted.
        /// </summary>
        warning,
        /// <summary>
        /// information warning style.
        /// </summary>
        information
    }
}
