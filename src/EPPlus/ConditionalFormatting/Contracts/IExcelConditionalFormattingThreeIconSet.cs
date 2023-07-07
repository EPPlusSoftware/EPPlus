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
  07/07/2023         EPPlus Software AB       Epplus 7
 *************************************************************************************************/

namespace OfficeOpenXml.ConditionalFormatting.Contracts
{
    /// <summary>
    /// IExcelConditionalFormattingThreeIconSet
    /// </summary>
    public interface IExcelConditionalFormattingThreeIconSet<T>
    : IExcelConditionalFormattingIconSetGroup<T>
	{
		#region Public Properties
    /// <summary>
    /// Icon1 (part of the 3, 4 or 5 Icon Set)
    /// </summary>
    ExcelConditionalFormattingIconDataBarValue Icon1 { get; }

    /// <summary>
    /// Icon2 (part of the 3, 4 or 5 Icon Set)
    /// </summary>
    ExcelConditionalFormattingIconDataBarValue Icon2 { get;  }

    /// <summary>
    /// Icon3 (part of the 3, 4 or 5 Icon Set)
    /// </summary>
    ExcelConditionalFormattingIconDataBarValue Icon3 { get; }
    #endregion Public Properties
	}
}