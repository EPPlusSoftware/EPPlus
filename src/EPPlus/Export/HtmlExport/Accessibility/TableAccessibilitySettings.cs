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
    /// This class controls how accessibility attributes will be set in the exported html.
    /// </summary>
    public class TableAccessibilitySettings
    {
        /// <summary>
        /// Reset all properties to their default value
        /// </summary>
        public void ResetToDefault()
        {
            AddAccessibilityAttributes = true;
            AriaLabel = string.Empty;
            AriaLabelledBy = string.Empty;
            TableRole = "table";
            TheadRole = "rowgroup";
            TbodyRole = "rowgroup";
            TfootRole = "rowgroup";
            TableHeaderCellRole = "columnheader";
        }

        /// <summary>
        /// Copies all properties from one instance to another.
        /// </summary>
        /// <param name="copy"></param>
        public void Copy(TableAccessibilitySettings copy)
        {
            AddAccessibilityAttributes = copy.AddAccessibilityAttributes;
            AriaLabel = copy.AriaLabel;
            AriaLabelledBy = copy.AriaLabelledBy;
            TableRole = copy.TableRole;
            TheadRole = copy.TheadRole;
            TbodyRole = copy.TbodyRole;
            TfootRole = copy.TfootRole;
            TableHeaderCellRole = copy.TableHeaderCellRole;
        }

        /// <summary>
        /// Controls whether accessibility attributes will be added to the html.
        /// </summary>
        public bool AddAccessibilityAttributes { get; set; }

        /// <summary>
        /// Value of the aria-label attribute
        /// </summary>
        public string AriaLabel { get; set; }

        /// <summary>
        /// Value of the aria-labelledby attribute
        /// </summary>
        public string AriaLabelledBy { get; set; }

        /// <summary>
        /// Value of the aria-describedby attribute
        /// </summary>
        public string AriaDescribedBy { get; set; }

        /// <summary>
        /// Value of the role attribute on the table element. Default value: table
        /// </summary>
        public string TableRole { get; set; }
        /// <summary>
        /// Value of the role attribute on the thead element. Default value: rowgroup
        /// </summary>
        public string TheadRole { get; set; }
        /// <summary>
        /// Value of the role attribute on the tbody element. Default value: rowgroup
        /// </summary>
        public string TbodyRole { get; set; }
        /// <summary>
        /// Value of the role attribute on the tfoot element. Default value: rowgroup
        /// </summary>
        public string TfootRole { get; set; }
        /// <summary>
        /// Value of the role attribute on the thead.tr.td element. Default value: columnheader
        /// </summary>
        public string TableHeaderCellRole { get; set; }
    }
}
