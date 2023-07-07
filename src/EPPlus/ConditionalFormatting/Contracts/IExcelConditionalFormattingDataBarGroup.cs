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
using System.Drawing;
using OfficeOpenXml.Style.Dxf;

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
        /// If the databar should be a gradient. True by default
        /// </summary>
        bool Gradient { get; set; }

        /// <summary>
        /// Wheter there is a border colour or not. 
        /// False by default. Is set to true if BorderColor or NegativeBorderColour is set
        /// </summary>
        bool Border { get; set; }

        /// <summary>
        /// Wheter negative and positive values should have the same colour. 
        /// False by default. Is set to true if NegativeFillColor is set.
        /// </summary>
        bool NegativeBarColorSameAsPositive { get; set; }

        /// <summary>
        /// Wheter negative and positive values should have the same border colour. 
        /// False by default. Is set to true if NegativeBorderColor is set.
        /// </summary>
        bool NegativeBarBorderColorSameAsPositive { get; set; }

        /// <summary>
        /// What position the axis between positive and negative values is to be put at.
        /// </summary>
        eExcelDatabarAxisPosition AxisPosition { get; set; }

        /// <summary>
        /// Databar Low Value
        /// </summary>
        ExcelConditionalFormattingIconDataBarValue LowValue { get;  }

        /// <summary>
        /// Databar High Value
        /// </summary>
        ExcelConditionalFormattingIconDataBarValue HighValue { get; }
        /// <summary>
        /// The color of the databar. ShortHand for FillColor.Color
        /// </summary>
        Color Color { get; set;}

        /// <summary>
        /// Fill color of Databar
        /// </summary>
        ExcelDxfColor FillColor { get; set; }
        /// <summary>
        /// Border color of databar. 
        /// Setting any property sets Border to true
        /// </summary>
        ExcelDxfColor BorderColor { get; set; }
        /// <summary>
        /// Fill color for negative values
        /// Setting any property sets NegativeBarColorSameAsPositive to false
        /// </summary>
        ExcelDxfColor NegativeFillColor { get; set; }
        /// <summary>
        /// Border color for negative values
        /// Setting any property sets NegativeBarBorderColorSameAsPositive to false
        /// </summary>
        ExcelDxfColor NegativeBorderColor { get; set; }
        /// <summary>
        /// Color of the axis between negative and positive values
        /// </summary>
        ExcelDxfColor AxisColor { get; set; }
        #endregion Public Properties
    }
}