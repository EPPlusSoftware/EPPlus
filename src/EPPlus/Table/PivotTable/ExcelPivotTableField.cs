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
using System.Xml;
using System.Globalization;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System.Linq;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Drawing;
using System.Text;
using System.Collections;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A pivot table field.
    /// </summary>
    public class ExcelPivotTableField : XmlHelper
    {
        internal ExcelPivotTable _table;
        internal ExcelPivotTableCacheField _cacheField = null;
        internal ExcelPivotTableField(XmlNamespaceManager ns, XmlNode topNode, ExcelPivotTable table, int index, int baseIndex) :
            base(ns, topNode)
        {
            Index = index;
            BaseIndex = baseIndex;
            _table = table;            
        }
        /// <summary>
        /// The index of the pivot table field
        /// </summary>
        public int Index
        {
            get;
            set;
        }
        /// <summary>
        /// The base line index of the pivot table field
        /// </summary>
        internal int BaseIndex
        {
            get;
            set;
        }
        /// <summary>
        /// Name of the field
        /// </summary>
        public string Name
        {
            get
            {
                string v = GetXmlNodeString("@name");
                if (v == "")
                {
                    return _cacheField.Name;
                }
                else
                {
                    return v;
                }
            }
            set
            {
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// Compact mode
        /// </summary>
        public bool Compact
        {
            get
            {
                return GetXmlNodeBool("@compact");
            }
            set
            {
                SetXmlNodeBool("@compact", value);
            }
        }
        /// <summary>
        /// A boolean that indicates whether the items in this field should be shown in Outline form
        /// </summary>
        public bool Outline
        {
            get
            {
                return GetXmlNodeBool("@outline");
            }
            set
            {
                SetXmlNodeBool("@outline", value);
            }
        }
        /// <summary>
        /// The custom text that is displayed for the subtotals label
        /// </summary>
        public bool SubtotalTop
        {
            get
            {
                return GetXmlNodeBool("@subtotalTop");
            }
            set
            {
                SetXmlNodeBool("@subtotalTop", value);
            }
        }
        /// <summary>
        /// Indicates whether the field can have multiple items selected in the page field
        /// </summary>
        public bool MultipleItemSelectionAllowed
        {
            get
            {
                return GetXmlNodeBool("@multipleItemSelectionAllowed");
            }
            set
            {
                SetXmlNodeBool("@multipleItemSelectionAllowed", value);
            }
        }
        #region Show properties
        /// <summary>
        /// Indicates whether to show all items for this field
        /// </summary>
        public bool ShowAll
        {
            get
            {
                return GetXmlNodeBool("@showAll");
            }
            set
            {
                SetXmlNodeBool("@showAll", value);
            }
        }
        /// <summary>
        /// Indicates whether to hide drop down buttons on PivotField headers
        /// </summary>
        public bool ShowDropDowns
        {
            get
            {
                return GetXmlNodeBool("@showDropDowns");
            }
            set
            {
                SetXmlNodeBool("@showDropDowns", value);
            }
        }
        /// <summary>
        /// Indicates whether this hierarchy is omitted from the field list
        /// </summary>
        public bool ShowInFieldList
        {
            get
            {
                return GetXmlNodeBool("@showInFieldList");
            }
            set
            {
                SetXmlNodeBool("@showInFieldList", value);
            }
        }
        /// <summary>
        /// Indicates whether to show the property as a member caption
        /// </summary>
        public bool ShowAsCaption
        {
            get
            {
                return GetXmlNodeBool("@showPropAsCaption");
            }
            set
            {
                SetXmlNodeBool("@showPropAsCaption", value);
            }
        }
        /// <summary>
        /// Indicates whether to show the member property value in a PivotTable cell
        /// </summary>
        public bool ShowMemberPropertyInCell
        {
            get
            {
                return GetXmlNodeBool("@showPropCell");
            }
            set
            {
                SetXmlNodeBool("@showPropCell", value);
            }
        }
        /// <summary>
        /// Indicates whether to show the member property value in a tooltip on the appropriate PivotTable cells
        /// </summary>
        public bool ShowMemberPropertyToolTip
        {
            get
            {
                return GetXmlNodeBool("@showPropTip");
            }
            set
            {
                SetXmlNodeBool("@showPropTip", value);
            }
        }
        #endregion
        /// <summary>
        /// The type of sort that is applied to this field
        /// </summary>
        public eSortType Sort
        {
            get
            {
                string v = GetXmlNodeString("@sortType");
                return v == "" ? eSortType.None : (eSortType)Enum.Parse(typeof(eSortType), v, true);
            }
            set
            {
                if (value == eSortType.None)
                {
                    DeleteNode("@sortType");
                }
                else
                {
                    SetXmlNodeString("@sortType", value.ToString().ToLower(CultureInfo.InvariantCulture));
                }
            }
        }
        /// <summary>
        /// A boolean that indicates whether manual filter is in inclusive mode
        /// </summary>
        public bool IncludeNewItemsInFilter
        {
            get
            {
                return GetXmlNodeBool("@includeNewItemsInFilter");
            }
            set
            {
                SetXmlNodeBool("@includeNewItemsInFilter", value);
            }
        }
        /// <summary>
        /// Enumeration of the different subtotal operations that can be applied to page, row or column fields
        /// </summary>
        public eSubTotalFunctions SubTotalFunctions
        {
            get
            {
                eSubTotalFunctions ret = 0;
                XmlNodeList nl = TopNode.SelectNodes("d:items/d:item/@t", NameSpaceManager);
                if (nl.Count == 0) return eSubTotalFunctions.None;
                foreach (XmlAttribute item in nl)
                {
                    try
                    {
                        ret |= (eSubTotalFunctions)Enum.Parse(typeof(eSubTotalFunctions), item.Value, true);
                    }
                    catch (ArgumentException ex)
                    {
                        throw new ArgumentException("Unable to parse value of " + item.Value + " to a valid pivot table subtotal function", ex);
                    }
                }
                return ret;
            }
            set
            {
                if ((value & eSubTotalFunctions.None) == eSubTotalFunctions.None && (value != eSubTotalFunctions.None))
                {
                    throw (new ArgumentException("Value None can not be combined with other values."));
                }
                if ((value & eSubTotalFunctions.Default) == eSubTotalFunctions.Default && (value != eSubTotalFunctions.Default))
                {
                    throw (new ArgumentException("Value Default can not be combined with other values."));
                }


                // remove old attribute                 
                XmlNodeList nl = TopNode.SelectNodes("d:items/d:item/@t", NameSpaceManager);
                if (nl.Count > 0)
                {
                    foreach (XmlAttribute item in nl)
                    {
                        DeleteNode("@" + item.Value + "Subtotal");
                        item.OwnerElement.ParentNode.RemoveChild(item.OwnerElement);
                    }
                }


                if (value == eSubTotalFunctions.None)
                {
                    // for no subtotals, set defaultSubtotal to off
                    SetXmlNodeBool("@defaultSubtotal", false);
                    //TopNode.InnerXml = "<items count=\"1\"><item x=\"0\"/></items>";
                    //_cacheFieldHelper.TopNode.InnerXml = "<sharedItems count=\"1\"><m/></sharedItems>";
                }
                else
                {
                    string innerXml = "";
                    int count = 0;
                    foreach (eSubTotalFunctions e in Enum.GetValues(typeof(eSubTotalFunctions)))
                    {
                        if ((value & e) == e)
                        {
                            var newTotalType = e.ToString();
                            var totalType = char.ToLowerInvariant(newTotalType[0]) + newTotalType.Substring(1);
                            // add new attribute
                            SetXmlNodeBool("@" + totalType + "Subtotal", true);
                            innerXml += "<item t=\"" + totalType + "\" />";
                            count++;
                        }
                    }
                    TopNode.InnerXml = string.Format("<items count=\"{0}\">{1}</items>", count, innerXml);
                }
            }
        }
        /// <summary>
        /// Type of axis
        /// </summary>
        public ePivotFieldAxis Axis
        {
            get
            {
                switch (GetXmlNodeString("@axis"))
                {
                    case "axisRow":
                        return ePivotFieldAxis.Row;
                    case "axisCol":
                        return ePivotFieldAxis.Column;
                    case "axisPage":
                        return ePivotFieldAxis.Page;
                    case "axisValues":
                        return ePivotFieldAxis.Values;
                    default:
                        return ePivotFieldAxis.None;
                }
            }
            internal set
            {
                switch (value)
                {
                    case ePivotFieldAxis.Row:
                        SetXmlNodeString("@axis", "axisRow");
                        break;
                    case ePivotFieldAxis.Column:
                        SetXmlNodeString("@axis", "axisCol");
                        break;
                    case ePivotFieldAxis.Values:
                        SetXmlNodeString("@axis", "axisValues");
                        break;
                    case ePivotFieldAxis.Page:
                        SetXmlNodeString("@axis", "axisPage");
                        break;
                    default:
                        DeleteNode("@axis");
                        break;
                }
            }
        }
        /// <summary>
        /// If the field is a row field
        /// </summary>
        public bool IsRowField
        {
            get
            {
                return (TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", Index), NameSpaceManager) != null);
            }
            internal set
            {
                if (value)
                {
                    var rowsNode = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);
                    if (rowsNode == null)
                    {
                        _table.CreateNode("d:rowFields");
                    }
                    rowsNode = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);

                    AppendField(rowsNode, Index, "field", "x");
                    if (Grouping == null)
                    {
                        if (BaseIndex == Index)
                        {
                            TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
                        }
                        else
                        {
                            TopNode.InnerXml = "<items count=\"0\"></items>";
                        }
                    }
                }
                else
                {
                    XmlElement node = TopNode.SelectSingleNode(string.Format("../../d:rowFields/d:field[@x={0}]", Index), NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        /// <summary>
        /// If the field is a column field
        /// </summary>
        public bool IsColumnField
        {
            get
            {
                return (TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", Index), NameSpaceManager) != null);
            }
            internal set
            {
                if (value)
                {
                    var columnsNode = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);
                    if (columnsNode == null)
                    {
                        _table.CreateNode("d:colFields");
                    }
                    columnsNode = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);

                    AppendField(columnsNode, Index, "field", "x");
                    if (BaseIndex == Index)
                    {
                        TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";
                    }
                    else
                    {
                        TopNode.InnerXml = "<items count=\"0\"></items>";
                    }
                }
                else
                {
                    XmlElement node = TopNode.SelectSingleNode(string.Format("../../d:colFields/d:field[@x={0}]", Index), NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        /// <summary>
        /// If the field is a datafield
        /// </summary>
        public bool IsDataField
        {
            get
            {
                return GetXmlNodeBool("@dataField", false);
            }
        }
        /// <summary>
        /// If the field is a page field.
        /// </summary>
        public bool IsPageField
        {
            get
            {
                return (Axis == ePivotFieldAxis.Page);
            }
            internal set
            {
                if (value)
                {
                    var dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
                    if (dataFieldsNode == null)
                    {
                        _table.CreateNode("d:pageFields");
                        dataFieldsNode = TopNode.SelectSingleNode("../../d:pageFields", NameSpaceManager);
                    }

                    TopNode.InnerXml = "<items count=\"1\"><item t=\"default\" /></items>";

                    XmlElement node = AppendField(dataFieldsNode, Index, "pageField", "fld");
                    _pageFieldSettings = new ExcelPivotTablePageFieldSettings(NameSpaceManager, node, this, Index);
                }
                else
                {
                    _pageFieldSettings = null;
                    XmlElement node = TopNode.SelectSingleNode(string.Format("../../d:pageFields/d:pageField[@fld={0}]", Index), NameSpaceManager) as XmlElement;
                    if (node != null)
                    {
                        node.ParentNode.RemoveChild(node);
                    }
                }
            }
        }
        //public ExcelPivotGrouping DateGrouping
        //{

        //}
        internal ExcelPivotTablePageFieldSettings _pageFieldSettings = null;
        /// <summary>
        /// Page field settings
        /// </summary>
        public ExcelPivotTablePageFieldSettings PageFieldSettings
        {
            get
            {
                return _pageFieldSettings;
            }
        }
        /// <summary>
        /// Date group by
        /// </summary>
        internal eDateGroupBy DateGrouping
        {
            get;
            set;
        }
        ExcelPivotTableFieldGroup _grouping = null;
        /// <summary>
        /// Grouping settings. 
        /// Null if the field has no grouping otherwise ExcelPivotTableFieldDateGroup or ExcelPivotTableFieldNumericGroup.
        /// </summary>        
        public ExcelPivotTableFieldGroup Grouping
        {
            get
            {
                return _grouping;
            }
        }
        #region Private & internal Methods
        internal XmlElement AppendField(XmlNode rowsNode, int index, string fieldNodeText, string indexAttrText)
        {
            XmlElement prevField = null, newElement;
            foreach (XmlElement field in rowsNode.ChildNodes)
            {
                string x = field.GetAttribute(indexAttrText);
                int fieldIndex;
                if (int.TryParse(x, out fieldIndex))
                {
                    if (fieldIndex == index)    //Row already exists
                    {
                        return field;
                    }
                }
                prevField = field;
            }
            newElement = rowsNode.OwnerDocument.CreateElement(fieldNodeText, ExcelPackage.schemaMain);
            newElement.SetAttribute(indexAttrText, index.ToString());
            rowsNode.InsertAfter(newElement, prevField);

            return newElement;
        }
        #endregion
        internal ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> _items = null;
        /// <summary>
        /// Pivottable field Items. Used for grouping.
        /// </summary>
        public ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem> Items
        {
            get
            {
                if (_items == null)
                {
                    _items = new ExcelPivotTableFieldCollectionBase<ExcelPivotTableFieldItem>(_table, Index);
                    if (_grouping == null)
                    {
                        var sh = _table.CacheDefinition._cacheReference.Fields[Index].SharedItems;
                        foreach (XmlElement node in TopNode.SelectNodes("d:items//d:item", NameSpaceManager))
                        {
                            var item = new ExcelPivotTableFieldItem(node);
                            if (item.X >= 0)
                            {
                                item.Value = sh[item.X];
                            }
                            _items.AddInternal(item);
                        }
                    }
                }
                return _items;
            }
        }
        public ExcelPivotTableCacheField Cache
        {
            get
            {
                return _table.CacheDefinition._cacheReference.Fields[Index];
            }
        }
        public ExcelPivotTableSlicer Slicer { get; internal set; }

        /// <summary>
        /// Add numberic grouping to the field
        /// </summary>
        /// <param name="Start">Start value</param>
        /// <param name="End">End value</param>
        /// <param name="Interval">Interval</param>
        public void AddNumericGrouping(double Start, double End, double Interval)
        {
            ValidateGrouping();
            _cacheField.SetNumericGroup(BaseIndex, Start, End, Interval);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="groupBy">Group by</param>
        public void AddDateGrouping(eDateGroupBy groupBy)
        {
            AddDateGrouping(groupBy, DateTime.MinValue, DateTime.MaxValue, 1);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="groupBy">Group by</param>
        /// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
        /// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
        public void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate)
        {
            AddDateGrouping(groupBy, startDate, endDate, 1);
        }
        /// <summary>
        /// Add a date grouping on this field.
        /// </summary>
        /// <param name="days">Number of days when grouping on days</param>
        /// <param name="startDate">Fixed start date. Use DateTime.MinValue for auto</param>
        /// <param name="endDate">Fixed end date. Use DateTime.MaxValue for auto</param>
        public void AddDateGrouping(int days, DateTime startDate, DateTime endDate)
        {
            AddDateGrouping(eDateGroupBy.Days, startDate, endDate, days);
        }
        private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField)
        {
            return AddField(groupBy, startDate, endDate, ref firstField, 1);
        }
        private ExcelPivotTableField AddField(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, ref bool firstField, int interval)
        {
            if (firstField == false)
            {
                //Pivot field
                var topNode = _table.PivotTableXml.SelectSingleNode("//d:pivotFields", _table.NameSpaceManager);
                var fieldNode = _table.PivotTableXml.CreateElement("pivotField", ExcelPackage.schemaMain);
                fieldNode.SetAttribute("compact", "0");
                fieldNode.SetAttribute("outline", "0");
                fieldNode.SetAttribute("showAll", "0");
                fieldNode.SetAttribute("defaultSubtotal", "0");
                topNode.AppendChild(fieldNode);

                var field = new ExcelPivotTableField(_table.NameSpaceManager, fieldNode, _table, _table.Fields.Count, Index);
                field.DateGrouping = groupBy;

                XmlNode rowColFields;
                if (IsRowField)
                {
                    rowColFields = TopNode.SelectSingleNode("../../d:rowFields", NameSpaceManager);
                }
                else
                {
                    rowColFields = TopNode.SelectSingleNode("../../d:colFields", NameSpaceManager);
                }

                int fieldIndex, index = 0;
                foreach (XmlElement rowfield in rowColFields.ChildNodes)
                {
                    if (int.TryParse(rowfield.GetAttribute("x"), out fieldIndex))
                    {
                        if (_table.Fields[fieldIndex].BaseIndex == BaseIndex)
                        {
                            var newElement = rowColFields.OwnerDocument.CreateElement("field", ExcelPackage.schemaMain);
                            newElement.SetAttribute("x", field.Index.ToString());
                            rowColFields.InsertBefore(newElement, rowfield);
                            break;
                        }
                    }
                    index++;
                }

                if (IsRowField)
                {
                    _table.RowFields.Insert(field, index);
                }
                else
                {
                    _table.ColumnFields.Insert(field, index);
                }

                field._cacheField = _table.CacheDefinition._cacheReference.AddDateGroupField(field, startDate, endDate, interval);

                _table.Fields.AddInternal(field);

                return field;
            }
            else
            {
                firstField = false;
                DateGrouping = groupBy;
                Compact = false;
                _cacheField.SetDateGroup(this, startDate, endDate, interval);
                return this;
            }
        }
        private void AddDateGrouping(eDateGroupBy groupBy, DateTime startDate, DateTime endDate, int groupInterval)
        {
            if (groupInterval < 1 || groupInterval >= Int16.MaxValue)
            {
                throw (new ArgumentOutOfRangeException("Group interval is out of range"));
            }
            if (groupInterval > 1 && groupBy != eDateGroupBy.Days)
            {
                throw (new ArgumentException("Group interval is can only be used when groupBy is Days"));
            }
            ValidateGrouping();

            bool firstField = true;
            int g = 0;
            //Seconds
            if ((groupBy & eDateGroupBy.Seconds) == eDateGroupBy.Seconds)
            {
                AddField(eDateGroupBy.Seconds, startDate, endDate, ref firstField);
                g++;
            }
            //Minutes
            if ((groupBy & eDateGroupBy.Minutes) == eDateGroupBy.Minutes)
            {
                AddField(eDateGroupBy.Minutes, startDate, endDate, ref firstField);
                g++;
            }
            //Hours
            if ((groupBy & eDateGroupBy.Hours) == eDateGroupBy.Hours)
            {
                AddField(eDateGroupBy.Hours, startDate, endDate, ref firstField);
                g++;
            }
            //Days
            if ((groupBy & eDateGroupBy.Days) == eDateGroupBy.Days)
            {
                AddField(eDateGroupBy.Days, startDate, endDate, ref firstField, groupInterval);
                g++;
            }
            //Month
            if ((groupBy & eDateGroupBy.Months) == eDateGroupBy.Months)
            {
                AddField(eDateGroupBy.Months, startDate, endDate, ref firstField);
                g++;
            }
            //Quarters
            if ((groupBy & eDateGroupBy.Quarters) == eDateGroupBy.Quarters)
            {
                AddField(eDateGroupBy.Quarters, startDate, endDate, ref firstField);
                g++;
            }
            //Years
            if ((groupBy & eDateGroupBy.Years) == eDateGroupBy.Years)
            {
                AddField(eDateGroupBy.Years, startDate, endDate, ref firstField);
                g++;
            }

            if (g>0) SetXmlNodeString("d:fieldGroup/@par", (g - 1).ToString());
            if (groupInterval != 1)
            {
                SetXmlNodeString("d:fieldGroup/d:rangePr/@groupInterval", groupInterval.ToString());
            }
            else
            {
                DeleteNode("d:fieldGroup/d:rangePr/@groupInterval");
            }
            _items = null;
        }

        private void ValidateGrouping()
        {
            if (!(IsColumnField || IsRowField))
            {
                throw (new Exception("Field must be a row or column field"));
            }
            foreach (var field in _table.Fields)
            {
                if (field.Grouping != null)
                {
                    throw (new Exception("Grouping already exists"));
                }
            }
        }
        internal string SaveToXml()
        {
            var sb = new StringBuilder();
            if (Axis != ePivotFieldAxis.Values && _table.CacheDefinition._cacheReference.Fields[Index].Grouping==null)
            {
                var cacheLookup = _table.CacheDefinition._cacheReference.Fields[Index]._cacheLookup;
                if(cacheLookup.Count==0)
                {
                    DeleteNode("d:items");       //Creates or return the existing node
                }
                else
                {
                    Items.Refresh();
                    foreach (var item in Items)
                    {
                        if(item.Type==eItemType.Data && cacheLookup.TryGetValue(item.Value, out int x))
                        {
                            item.X = cacheLookup[item.Value];
                        }
                        else
                        {
                            item.X = -1;
                        }
                        item.GetXmlString(sb);
                    }
                    var node = (XmlElement)CreateNode("d:items");       //Creates or return the existing node
                    node.InnerXml = sb.ToString();
                    node.SetAttribute("count", Items.Count.ToString());
                }
            }
            return sb.ToString();
        }
    }
}
