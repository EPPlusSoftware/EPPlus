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
namespace OfficeOpenXml.Interfaces.SensitivityLabels;

/// <summary>
/// Interfaced applied to update properties on the <see cref="IExcelSensibilityLabel"/>
/// </summary>
public interface IExcelSensibilityLabelUpdate
{
    /// <summary>
    /// Updates the properties on the label.
    /// </summary>
    /// <param name="name">The name to update</param>
    /// <param name="tooltip">The tooltip to update</param>
    /// <param name="description">The description to update</param>
    /// <param name="color">The color to update</param>
    /// <param name="parent">The parent to update</param>
    public void Update(string name, string tooltip, string description, string color, IExcelSensibilityLabel parent);
}
#endif