/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  09/18/2024         EPPlus Software AB       EPPlus 8
 *************************************************************************************************/
#if(!NET35)
using System.Collections.Generic;

namespace OfficeOpenXml.Interfaces.SensitivityLabels;

/// <summary>
/// Represents a Sensibility Label that can be applied to a package.
/// </summary>
public interface IExcelSensibilityLabel 
{
    /// <summary>
    /// The Id of the sensibility label
    /// </summary>
    public string Id { get;  }
    /// <summary>
    /// The name of the sensibility label. 
    /// This property will be null if it's not updated by the SensitivityLabelHandler. 
    /// <see cref="ISensitivityLabelHandler.UpdateLabelList(IEnumerable{IExcelSensibilityLabel}, string)"/>
    /// </summary>
    public string Name { get; }
    /// <summary>
    /// The description of the sensibility label. 
    /// This property will be null if it's not updated by the SensitivityLabelHandler. 
    /// <see cref="ISensitivityLabelHandler.UpdateLabelList(IEnumerable{IExcelSensibilityLabel}, string)"/>
    /// </summary>
    public string Description { get; }
    /// <summary>
    /// The tooltip of the sensibility label. 
    /// This property will be null if it's not updated by the SensitivityLabelHandler. 
    /// <see cref="ISensitivityLabelHandler.UpdateLabelList(IEnumerable{IExcelSensibilityLabel}, string)"/>
    /// </summary>
    public string Tooltip { get; }
    /// <summary>
    /// The parent of the sensibility label. 
    /// This property will be null if it's not updated by the SensitivityLabelHandler. 
    /// <see cref="ISensitivityLabelHandler.UpdateLabelList(IEnumerable{IExcelSensibilityLabel}, string)"/>
    /// </summary>
    public IExcelSensibilityLabel Parent { get; }
    /// <summary>
    /// The color of the sensibility label. 
    /// This property will be null if it's not updated by the SensitivityLabelHandler. 
    /// <see cref="ISensitivityLabelHandler.UpdateLabelList(IEnumerable{IExcelSensibilityLabel}, string)"/>
    /// </summary>
    public string Color { get; }
    /// <summary>
    /// If the sensibility label is enabled. Only on sensibility label can be enabled at once.
    /// </summary>
    public bool Enabled { get; }
    /// <summary>
    /// If the sensibility label has been removed and is no longer active.
    /// </summary>
    public bool Removed { get; }
    /// <summary>
    /// The Site/Tenent id for this label.
    /// </summary>
    public string SiteId { get; }
    /// <summary>
    /// How the label has been applied.
    /// </summary>
    public eMethod Method { get; }
    /// <summary>
    /// Information about encryption and content applied.
    /// </summary>
    public eContentBits ContentBits { get; }
}
#endif