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
	/// IExcelConditionalFormattingIconSetGroup
	/// </summary>
	public interface IExcelConditionalFormattingIconSetGroup<T>
		: IExcelConditionalFormattingRule
	{
		#region Public Properties
    /// <summary>
    /// Reverse
    /// </summary>
    bool Reverse { get; set; }

    /// <summary>
    /// ShowValue
    /// </summary>
    bool ShowValue { get; set; }

    /// <summary>
    /// True if percent based
    /// </summary>
    bool IconSetPercent { get; set; }

    /// <summary>
    /// True if the Iconset has custom icons
    /// </summary>
    bool Custom { get; }

    /// <summary>
    /// IconSet (3, 4 or 5 IconSet)
    /// </summary>
    T IconSet { get; set; }
    #endregion Public Properties
	}
}