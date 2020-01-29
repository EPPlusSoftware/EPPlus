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
	/// IExcelConditionalFormattingTwoColorScale
	/// </summary>
	public interface IExcelConditionalFormattingTwoColorScale
		: IExcelConditionalFormattingColorScaleGroup
	{
		#region Public Properties
    /// <summary>
    /// Two Color Scale Low Value
    /// </summary>
    ExcelConditionalFormattingColorScaleValue LowValue { get; set; }

    /// <summary>
    /// Two Color Scale High Value
    /// </summary>
    ExcelConditionalFormattingColorScaleValue HighValue { get; set; }
    #endregion Public Properties
	}
}