/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  05/16/2020         EPPlus Software AB           ExcelTable Html Export
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.HtmlExport
{
    /// <summary>
    /// When exporting multiple ranges from the same workbook, this class can be used
    /// to override certain properties of the settings.
    /// </summary>
    public class ExcelHtmlOverrideExportSettings
    {
        /// <summary>
        /// Html id of the exported table.
        /// </summary>
        public string TableId { get; set; }

        /// <summary>
        /// Use this property to set additional class names that will be set on the exported html-table.
        /// </summary>
        public List<string> AdditionalTableClassNames
        {
            get;
            protected internal set;
        } = new List<string>();
    }
}
