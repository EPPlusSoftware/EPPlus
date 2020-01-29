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
using System.Drawing;

namespace OfficeOpenXml.ConditionalFormatting.Contracts
{
	/// <summary>
	/// IExcelConditionalFormattingDataBar
	/// </summary>
	public interface IExcelConditionalFormattingDataBarGroup
        : IExcelConditionalFormattingRule
	{
		#region Public Properties
        /// <summary>
        /// ShowValue
        /// </summary>
        bool ShowValue { get; set; }
        /// <summary>
        /// Databar Low Value
        /// </summary>
        ExcelConditionalFormattingIconDataBarValue LowValue { get;  }

        /// <summary>
        /// Databar High Value
        /// </summary>
        ExcelConditionalFormattingIconDataBarValue HighValue { get; }
        /// <summary>
        /// The color of the databar
        /// </summary>
        Color Color { get; set;}
        #endregion Public Properties
	}
}