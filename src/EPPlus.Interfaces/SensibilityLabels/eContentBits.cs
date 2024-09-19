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
using System;
namespace OfficeOpenXml.Interfaces.SensitivityLabels;
[Flags]
public enum eContentBits
{
    /// <summary>
    /// No content information
    /// </summary>
    None = 0,
    /// <summary>
    /// Content to the header will be applied.
    /// </summary>
    Header = 1,
    /// <summary>
    /// Content to the footer will be applied.
    /// </summary>
    Footer = 2,
    /// <summary>
    /// A watermark will be applied.
    /// </summary>
    Watermark = 4,
    /// <summary>
    /// The label will encrypt the package.
    /// </summary>
    Encryption = 8
}
#endif
