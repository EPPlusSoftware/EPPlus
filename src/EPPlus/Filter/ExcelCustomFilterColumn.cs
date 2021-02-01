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
using System.Xml;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Utils.Extensions;
namespace OfficeOpenXml.Filter
{
    /// <summary>
    /// Represents a custom filter column
    /// </summary>
    public class ExcelCustomFilterColumn : ExcelFilterColumn
    {
        internal ExcelCustomFilterColumn(XmlNamespaceManager namespaceManager, XmlNode topNode) : base(namespaceManager, topNode)
        {
            Filters = new ExcelFilterCollection<ExcelFilterCustomItem>();
            LoadFilters(topNode);
        }
        private void LoadFilters(XmlNode topNode)
        {
            foreach (XmlElement node in topNode.FirstChild.ChildNodes)
            {
                if(node.HasAttribute("and"))
                {
                    And = node.GetAttribute("and") == "1";
                }

                switch (node.LocalName)
                {
                    case "customFilter":
                        eFilterOperator filterOperator;
                        try
                        {
                            filterOperator = (eFilterOperator)Enum.Parse(typeof(eFilterOperator), node.Attributes["operator"].Value, true);
                        }
                        catch
                        {
                            filterOperator = eFilterOperator.Equal;
                        }
                        Filters.Add(new ExcelFilterCustomItem(node.Attributes["val"].Value, filterOperator));
                        break;
                }
            }
        }
        bool _isNumericFilterSet;
        bool _isNumericFilter;
        /// <summary>
        /// If true filter is numeric otherwise it's textual.
        /// If this property is not set, the value is set from the first value in column of the filtered range
        /// </summary>
        public bool IsNumericFilter
        {
            get
            {
                return _isNumericFilter;
            }
            set
            {
                _isNumericFilter = value;
                _isNumericFilterSet = true;
            }
        }
        /// <summary>
        /// Flag indicating whether the two criteria have an "and" relationship. true indicates "and", false indicates "or".
        /// </summary>
        public bool And
        {
            get;
            set;
        }
        /// <summary>
        /// The filters to apply
        /// </summary>
        public ExcelFilterCollection<ExcelFilterCustomItem> Filters { get; set; }

        internal override bool Match(object value, string valueText)
        {
            if(!_isNumericFilterSet)
            {
                IsNumericFilter = Utils.ConvertUtil.IsNumericOrDate(value);
            }
            var match = true;
            foreach(var filter in Filters)
            {
                if(IsNumericFilter)
                {
                    match = MatchByOperatorNumeric(value, filter);
                }
                else
                {
                    match = MatchByOperatorText(valueText, filter);
                }

                if (match == false && And)
                    return false;
                else if (match && And == false)
                    return true;
            }
            return match;
        }

        private bool MatchByOperatorNumeric(object value, ExcelFilterCustomItem filter)
        {
            if (filter.Operator == null)
            {
                return filter._valueDouble.Equals(Utils.ConvertUtil.GetValueDouble(value));
            }
            else
            {
                switch (filter.Operator.Value)
                {
                    case eFilterOperator.Equal:
                        return filter._valueDouble.Equals(Utils.ConvertUtil.GetValueDouble(value));
                    case eFilterOperator.NotEqual:
                        return !filter._valueDouble.Equals(Utils.ConvertUtil.GetValueDouble(value));
                    case eFilterOperator.GreaterThan:
                        return Utils.ConvertUtil.GetValueDouble(value) > filter._valueDouble;
                    case eFilterOperator.GreaterThanOrEqual:
                        return Utils.ConvertUtil.GetValueDouble(value) >= filter._valueDouble;
                    case eFilterOperator.LessThan:
                        return Utils.ConvertUtil.GetValueDouble(value) < filter._valueDouble;
                    case eFilterOperator.LessThanOrEqual:
                        return Utils.ConvertUtil.GetValueDouble(value) <= filter._valueDouble;
                    default:
                        throw (new ArgumentException($"Unhandled filter operator {filter.Operator}"));
                }
            }
        }
        private static bool MatchByOperatorText(object value, ExcelFilterCustomItem filter)
        {
            if (filter.Operator == null)
            {
                return FilterWildCardMatcher.Match(value.ToString(), filter.Value);
            }
            else
            {
                switch (filter.Operator.Value)
                {
                    case eFilterOperator.Equal:
                        return FilterWildCardMatcher.Match(value.ToString(), filter.Value);
                    case eFilterOperator.NotEqual:
                        return !FilterWildCardMatcher.Match(value.ToString(), filter.Value);
                    case eFilterOperator.GreaterThan:
                        return string.Compare(value.ToString(), filter.Value, StringComparison.CurrentCultureIgnoreCase) > 0;
                    case eFilterOperator.GreaterThanOrEqual:
                        return string.Compare(value.ToString(), filter.Value, StringComparison.CurrentCultureIgnoreCase) >= 0;
                    case eFilterOperator.LessThan:
                        return string.Compare(value.ToString(), filter.Value, StringComparison.CurrentCultureIgnoreCase) < 0;
                    case eFilterOperator.LessThanOrEqual:
                        return string.Compare(value.ToString(), filter.Value, StringComparison.CurrentCultureIgnoreCase) <= 0;
                    default:
                        throw (new ArgumentException($"Unhandled filter operator {filter.Operator}"));
                }
            }
        }

            internal override void Save()
            {
                var node = (XmlElement)CreateNode("d:customFilters");
                node.RemoveAll();
                if (And) node.SetAttribute("and", "1");
                foreach (var f in Filters)
                {
                    var e = TopNode.OwnerDocument.CreateElement("customFilter", ExcelPackage.schemaMain);
                    e.SetAttribute("val", f.Value);
                    if(f.Operator.HasValue) e.SetAttribute("operator", f.Operator.Value.ToEnumString());
                    node.AppendChild(e);
                }
            }
        }
}