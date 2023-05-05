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
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using OfficeOpenXml.Style.Dxf;

namespace OfficeOpenXml.ConditionalFormatting.Contracts
{
    /// <summary>
    /// Interface for conditional formatting rule
    /// </summary>
    public interface IExcelConditionalFormattingRule
    {
        /// <summary>
        /// The 'cfRule' XML node
        /// </summary>
        XmlNode Node { get; }

        /// <summary>
        /// The type of conditional formatting rule.
        /// </summary>
        eExcelConditionalFormattingRuleType Type { get; }

        /// <summary>
        /// <para>The range over which these conditional formatting rules apply.</para>
        /// </summary>
        ExcelAddress Address { get; set; }

        /// <summary>
        /// The priority of the rule. 
        /// A lower values are higher priority than higher values, where 1 is the highest priority.
        /// </summary>
        int Priority { get; set; }

        /// <summary>
        /// If this property is true, no rules with lower priority should be applied over this rule.
        /// </summary>
        bool StopIfTrue { get; set; }

        /// <summary>
        /// Gives access to the differencial styling (DXF) for the rule.
        /// </summary>
        ExcelDxfStyleConditionalFormatting Style { get; }

        /// <summary>
        /// Indicates that the conditional formatting is associated with a PivotTable
        /// </summary>
        bool PivotTable { get; set; }
        /// <summary>
        /// Type case propterty for the base class.
        /// </summary>
        ExcelConditionalFormattingAsType As { get; }
    }
}