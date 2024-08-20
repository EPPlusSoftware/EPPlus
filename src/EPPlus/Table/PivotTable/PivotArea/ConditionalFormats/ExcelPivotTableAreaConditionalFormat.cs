/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/28/2020         EPPlus Software AB       Pivot Table Styling - EPPlus 5.6
 *************************************************************************************************/
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.ConditionalFormatting.Contracts;
using OfficeOpenXml.Utils.Extensions;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Defines a pivot table area of selection used for styling.
    /// </summary>
    public class ExcelPivotTableAreaConditionalFormat : XmlHelper
    {
        ExcelConditionalFormattingCollection _conditionalFormattings;
        internal ExcelPivotTableAreaConditionalFormat(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt) :
            base(nsm, topNode)
        {
            _conditionalFormattings = pt.WorkSheet.ConditionalFormatting;
            foreach(var cf in _conditionalFormattings)
            {
                if(cf.Priority==Priority)
                {
                    ConditionalFormatting = cf;
                }
            }
            var node = CreateNode("d:pivotAreas");
            Areas = new ExcelPivotTableAreas(pt, node);
        }

        internal ExcelPivotTableAreaConditionalFormat(XmlNamespaceManager nsm, XmlNode topNode, ExcelPivotTable pt, eExcelConditionalFormattingRuleType type) :
            base(nsm, topNode)
        {
            _conditionalFormattings = pt.WorkSheet.ConditionalFormatting;

            ConditionalFormatting = _conditionalFormattings.AddRule(type, new ExcelAddress(pt.Address.Address), true);
            ConditionalFormatting.PivotTable = true;
            Priority = ConditionalFormatting.Priority;
            var node = CreateNode("d:pivotAreas");
            Areas = new ExcelPivotTableAreas(pt, node);
        }
        /// <summary>
        /// A collection of conditions for the conditional formats. Conditions can be set for specific row-, column- or data fields. Specify labels, data grand totals and more.
        /// </summary>
        public ExcelPivotTableAreas Areas
        {
            get;
        }

        IExcelConditionalFormattingRule _conditionalFormatting = null;
        /// <summary>
        /// Access to the style property for the pivot area
        /// </summary>
        public IExcelConditionalFormattingRule ConditionalFormatting
        { 
            get
            {
                if (_conditionalFormatting == null)
                {
                    _conditionalFormatting = _conditionalFormattings.GetByPriority(Priority);
                }
                return _conditionalFormatting;
            }
            internal set
            {
                _conditionalFormatting = value;
            }
        }
        /// <summary>
        /// The priority of the pivot table conditional formatting rule that should be matched in the worksheet.
        /// </summary>
        internal int Priority 
        { 
            get
            {
                return GetXmlNodeInt("@priority");                
            }
            private set
            {
                SetXmlNodeInt("@priority", value);
            }
        }
        /// <summary>
        /// The condition type of the pivot table conditional formatting rule. Default is None
        /// </summary>
        public ePivotTableConditionalFormattingConditionType Type
        {
            get
            {
                return GetXmlEnum("@type", ePivotTableConditionalFormattingConditionType.None);
            }
            set
            {
                SetXmlNodeString("@type", value.ToEnumString());
            }
        }
        /// <summary>
        /// The scope of the pivot table conditional formatting rule. Default is Selection.
        /// </summary>
        public ePivotTableConditionalFormattingConditionScope Scope
        {
            get
            {
                return GetXmlEnum("@scope", ePivotTableConditionalFormattingConditionScope.Selection);
            }
            set
            {
                SetXmlNodeString("@scope", value.ToEnumString());
            }
        }
    }
    /// <summary>
    /// Conditional Formatting Evaluation Type
    /// </summary>
    public enum ePivotTableConditionalFormattingConditionType
    {
        /// <summary>
        /// The conditional formatting is not evaluated
        /// </summary>
        None,
        /// <summary>
        /// The Top N conditional formatting is evaluated across the entire scope range.
        /// </summary>
        All,
        /// <summary>
        /// The Top N conditional formatting is evaluated for each row§.
        /// </summary>
        Row,
        /// <summary>
        /// The Top N conditional formatting is evaluated for each column.
        /// </summary>
        Column
    }
    /// <summary>
    /// The scope of the pivot table conditional formatting rule
    /// </summary>
    public enum ePivotTableConditionalFormattingConditionScope
    {
        /// <summary>
        /// The conditional formatting is applied to the selected data fields.
        /// </summary>
        Data,
        /// <summary>
        /// The conditional formatting is applied to the selected PivotTable field intersections.
        /// </summary>
        Field,
        /// <summary>
        /// The conditional formatting is applied to the selected data fields.
        /// </summary>
        Selection
    }
}

