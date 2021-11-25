/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/11/2021         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport.Accessibility
{
    /// <summary>
    /// This class contains settings for usage of accessibility/ARIA attributes in the exported html.
    /// </summary>
    public class AccessibilitySettings
    {
        internal AccessibilitySettings()
        {
            TableSettings.ResetToDefault();
        }

        /// <summary>
        /// Settings for a html table
        /// </summary>
        public TableAccessibilitySettings TableSettings { get; private set; } = new TableAccessibilitySettings();

    }
}
