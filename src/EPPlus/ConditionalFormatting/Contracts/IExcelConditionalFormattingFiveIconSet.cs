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
using System.Collections.Generic;
using System.Linq;
using System.Text;

using OfficeOpenXml.ConditionalFormatting;

namespace OfficeOpenXml.ConditionalFormatting.Contracts
{
  /// <summary>
  /// IExcelConditionalFormattingFiveIconSet
  /// </summary>eExcelconditionalFormatting4IconsSetType
    public interface IExcelConditionalFormattingFiveIconSet : IExcelConditionalFormattingFourIconSet<eExcelconditionalFormatting5IconsSetType>
  {
    #region Public Properties
    /// <summary>
    /// Icon5 (part of the 5 Icon Set)
    /// </summary>
      ExcelConditionalFormattingIconDataBarValue Icon5 { get; }
    #endregion Public Properties
  }
}