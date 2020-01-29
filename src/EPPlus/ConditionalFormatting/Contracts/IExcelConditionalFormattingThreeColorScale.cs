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
	/// IExcelConditionalFormattingThreeColorScale
	/// </summary>
	public interface IExcelConditionalFormattingThreeColorScale
    : IExcelConditionalFormattingTwoColorScale
	{
		#region Public Properties
		/// <summary>
		/// Three Color Scale Middle Value
		/// </summary>
		ExcelConditionalFormattingColorScaleValue MiddleValue { get; set; }
		#endregion Public Properties
	}
}