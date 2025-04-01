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
/// The method used to apply a sensibility label.
/// </summary>
public enum eMethod
{
    /// <summary>
    /// If the sensibility label is removed, the value should be empty.
    /// </summary>
    Empty,
    /// <summary>
    /// Use for any sensitivity label that was not directly applied by the user. This includes any default labels, automatically applied labels.
    /// </summary>
    Standard,
    /// <summary>
    /// Use for any sensitivity label that was directly applied by the user. This includes any manually applied sensitivity labels as well as recommended or mandatory labeling or any feature where the user decides which sensitivity label to apply.
    /// </summary>
    Privileged
}
#endif