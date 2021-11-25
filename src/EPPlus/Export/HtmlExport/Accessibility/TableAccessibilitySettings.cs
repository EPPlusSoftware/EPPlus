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
    public class TableAccessibilitySettings
    {
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

        public bool AddAccessibilityAttributes { get; set; }

        public string AriaLabel { get; set; }

        public string AriaLabelledBy { get; set; }

        public string AriaDescribedBy { get; set; }

        public string TableRole { get; set; }

        public string TheadRole { get; set; }

        public string TbodyRole { get; set; }

        public string TfootRole { get; set; }

        public string TableHeaderCellRole { get; set; }
    }
}
