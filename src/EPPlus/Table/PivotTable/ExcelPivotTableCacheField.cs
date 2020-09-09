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
using OfficeOpenXml.ConditionalFormatting;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    public class ExcelPivotTableCacheField  : XmlHelper
    {
        [Flags]
        private enum DataTypeFlags
        {
            Empty = 0x1,
            String = 0x2,
            Int = 0x4,
            Number = 0x8,
            DateTime = 0x10,
            Boolean = 0x20,
            Error = 0x30
        }
        private PivotTableCacheInternal _cache;
        internal ExcelPivotTableCacheField(XmlNamespaceManager nsm, XmlNode topNode,  PivotTableCacheInternal cache, int index) : base(nsm, topNode)
        {
            _cache = cache;
            Index = index;
            SetCacheFieldNode();
        }
        public int Index { get; set; }
        public string Name
        {
            get
            {
                return GetXmlNodeString("@name");
            }
            internal set
            {
                SetXmlNodeString("@name", value);
            }
        }
        /// <summary>
        /// A list of unique items for the field 
        /// </summary>
        public List<object> SharedItems
        {
            get;
            set;
        } = new List<object>();
        /// <summary>
        /// A list of group items, if the field has grouping.
        /// <seealso cref="Grouping"/>
        /// </summary>
        public List<object> GroupItems
        {
            get;
            set;
        } = new List<object>();

        internal Dictionary<object, int> _cacheLookup=null;
        public ExcelPivotTableSlicer Slicer { get; internal set; }
        public eDateGroupBy DateGrouping { get; private set; }
        public ExcelPivotTableFieldGroup Grouping { get; set; }

        internal void WriteSharedItems(XmlElement fieldNode, XmlNamespaceManager nsm)
        {
            var shNode = (XmlElement)fieldNode.SelectSingleNode("d:sharedItems", nsm);
            shNode.RemoveAll();

            var flags = GetFlags();

            _cacheLookup = new Dictionary<object, int>();
            if (IsRowColumnOrPage)
            {
                AppendSharedItems(shNode);
            }
            if (!HasOneValueOnly(flags) && flags!=(DataTypeFlags.Int| DataTypeFlags.Number) && SharedItems.Count>1)
            {
                if ((flags & DataTypeFlags.String) == DataTypeFlags.String)
                {
                    shNode.SetAttribute("containsSemiMixedTypes", "1");
                }
                else
                {
                    shNode.SetAttribute("containsMixedTypes", "1");
                }
                SetFlags(shNode, flags);
            }
            else
            {
                if ((flags & DataTypeFlags.String) != DataTypeFlags.String &&
                    (flags & DataTypeFlags.Empty) != DataTypeFlags.Empty)
                {
                    shNode.SetAttribute("containsSemiMixedTypes", "0");
                    shNode.SetAttribute("containsString", "0");
                }
                SetFlags(shNode, flags);
            }
        }
        internal bool IsRowColumnOrPage
        {
            get
            {
                foreach(var pt in _cache._pivotTables)
                {
                    if (Index < pt.Fields.Count)
                    {
                        var axis = pt.Fields[Index].Axis;
                        if (axis == ePivotFieldAxis.Column ||
                            axis == ePivotFieldAxis.Row ||
                            axis == ePivotFieldAxis.Page)
                        {
                            return true;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                return false;
            }
        }

        private void SetFlags(XmlElement shNode, DataTypeFlags flags)
        {
            if((flags & DataTypeFlags.DateTime) == DataTypeFlags.DateTime)
            {
                shNode.SetAttribute("containsDate", "1");
            }
            if ((flags & DataTypeFlags.Number) == DataTypeFlags.Number)
            {
                shNode.SetAttribute("containsNumber", "1");
            }
            if ((flags & DataTypeFlags.Int) == DataTypeFlags.Int)
            {
                shNode.SetAttribute("containsInteger", "1");
            }
            if ((flags & DataTypeFlags.Empty) == DataTypeFlags.Empty)
            {
                shNode.SetAttribute("containsBlank", "1");
            }
            shNode.SetAttribute("containsString", (flags & DataTypeFlags.String) == DataTypeFlags.String ? "1" : "0");
        }

        private bool HasOneValueOnly(DataTypeFlags flags)
        {
            foreach(DataTypeFlags v in Enum.GetValues(typeof(DataTypeFlags)))
            {
                if(flags==v)
                {
                    return true;
                }
            }
            return false;
        }

        private void AppendSharedItems(XmlElement shNode)
        {
            int index = 0;
            foreach (var si in SharedItems)
            {
                if (si == null)
                {
                    _cacheLookup.Add("", index++);  //Can't have null as key.
                    AppendItem(shNode, "m", null);
                }
                else
                {
                    _cacheLookup.Add(si, index++);
                    var t = si.GetType();
                    switch (Type.GetTypeCode(t))
                    {
                        case TypeCode.String:
                        case TypeCode.Char:
                            AppendItem(shNode, "s", si.ToString());
                            break;
                        case TypeCode.Byte:
                        case TypeCode.SByte:
                        case TypeCode.UInt16:
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                        case TypeCode.Int16:
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                        case TypeCode.Decimal:
                        case TypeCode.Double:
                        case TypeCode.Single:
                            AppendItem(shNode, "n", ConvertUtil.GetValueForXml(si, false));
                            break;
                        case TypeCode.DateTime:
                            AppendItem(shNode, "d", ((DateTime)si).ToString("s"));
                            break;
                        case TypeCode.Boolean:
                            AppendItem(shNode, "b", ConvertUtil.GetValueForXml(si, false));
                            break;
                        case TypeCode.Empty:
                            AppendItem(shNode, "m", null);
                            break;
                        default:
                            if (t == typeof(TimeSpan))
                            {
                                AppendItem(shNode, "d", ConvertUtil.GetValueForXml(si, false));
                            }
                            else if (t == typeof(ExcelErrorValue))
                            {
                                AppendItem(shNode, "e", si.ToString());
                            }
                            else
                            {
                                AppendItem(shNode, "s", si.ToString());
                            }
                            break;
                    }

                }
            }
        }

        private DataTypeFlags GetFlags()
        {
            DataTypeFlags flags = 0;
            foreach (var si in SharedItems)
            {
                if (si == null)
                {
                    flags |= DataTypeFlags.Empty;
                }
                else
                {
                    var t = si.GetType();
                    switch (Type.GetTypeCode(t))
                    {
                        case TypeCode.String:
                        case TypeCode.Char:
                            flags |= DataTypeFlags.String;
                            break;
                        case TypeCode.Byte:
                        case TypeCode.SByte:
                        case TypeCode.UInt16:
                        case TypeCode.UInt32:
                        case TypeCode.UInt64:
                        case TypeCode.Int16:
                        case TypeCode.Int32:
                        case TypeCode.Int64:
                            flags |= (DataTypeFlags.Number|DataTypeFlags.Int);
                            break;
                        case TypeCode.Decimal:
                        case TypeCode.Double:
                        case TypeCode.Single:
                            flags |= DataTypeFlags.Number;
                            break;
                        case TypeCode.DateTime:
                            flags |= DataTypeFlags.DateTime;
                            break;
                        case TypeCode.Boolean:
                            flags |= DataTypeFlags.Boolean;
                            break;
                        case TypeCode.Empty:
                            flags |= DataTypeFlags.Empty;
                            break;
                        default:
                            if (t == typeof(TimeSpan))
                            {
                                flags |= DataTypeFlags.DateTime;
                            }
                            else if(t==typeof(ExcelErrorValue))
                            {
                                flags |= DataTypeFlags.Error;
                            }
                            else
                            {
                                flags |= DataTypeFlags.String;
                            }
                            break;
                    }
                }
            }
            return flags;
        }
        private void AppendItem(XmlElement shNode, string elementName, string value)
        {
            var e = shNode.OwnerDocument.CreateElement(elementName, ExcelPackage.schemaMain);
            if (value != null)
            {
                e.SetAttribute("v", value);
            }
            shNode.AppendChild(e);
        }
        internal void SetCacheFieldNode()
        {
            var groupNode = GetNode("d:fieldGroup");
            if (groupNode != null)
            {
                var groupBy = groupNode.SelectSingleNode("d:rangePr/@groupBy", NameSpaceManager);
                if (groupBy == null)
                {
                    Grouping = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, TopNode);
                }
                else
                {
                    DateGrouping = (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), groupBy.Value, true);
                    Grouping = new ExcelPivotTableFieldDateGroup(NameSpaceManager, groupNode);
                }
                AddItems(GroupItems, groupNode.SelectSingleNode("d:groupItems", NameSpaceManager), true);
            }

            var si = GetNode("d:sharedItems");
            if (si != null)
            {
                AddItems(SharedItems, si, groupNode==null);
            }

        }

        private void AddItems(List<Object> items, XmlNode itemsNode, bool updateCacheLookup)
        {
            if (updateCacheLookup)
            {
                _cacheLookup = new Dictionary<object, int>();
            }

            foreach (XmlElement c in itemsNode.ChildNodes)
            {
                if (c.LocalName == "s")
                {
                    items.Add(c.Attributes["v"].Value);
                }
                else if (c.LocalName == "d")
                {
                    if (ConvertUtil.TryParseDateString(c.Attributes["v"].Value, out DateTime d))
                    {
                        items.Add(d);
                    }
                    else
                    {
                        items.Add(c.Attributes["v"].Value);
                    }
                }
                else if (c.LocalName == "n")
                {
                    if (ConvertUtil.TryParseNumericString(c.Attributes["v"].Value, out double num))
                    {
                        items.Add(num);
                    }
                    else
                    {
                        items.Add(c.Attributes["v"].Value);
                    }
                }
                else if (c.LocalName == "b")
                {
                    if (ConvertUtil.TryParseBooleanString(c.Attributes["v"].Value, out bool b))
                    {
                        items.Add(b);
                    }
                    else
                    {
                        items.Add(c.Attributes["v"].Value);
                    }
                }
                else if (c.LocalName == "e")
                {
                    if (ExcelErrorValue.Values.StringIsErrorValue(c.Attributes["v"].Value))
                    {
                        items.Add(ExcelErrorValue.Parse(c.Attributes["v"].Value));
                    }
                    else
                    {
                        items.Add(c.Attributes["v"].Value);
                    }
                }
                else
                {
                    items.Add(null);
                }
                if(updateCacheLookup)
                {
                    _cacheLookup.Add(items[items.Count-1]??"", items.Count - 1);
                }
            }
        }
        #region Grouping
        internal ExcelPivotTableFieldDateGroup SetDateGroup(ExcelPivotTableField field, DateTime StartDate, DateTime EndDate, int interval)
        {
            ExcelPivotTableFieldDateGroup group;
            group = new ExcelPivotTableFieldDateGroup(NameSpaceManager, TopNode);
            SetXmlNodeBool("d:sharedItems/@containsDate", true);
            SetXmlNodeBool("d:sharedItems/@containsNonDate", false);
            SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

            group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>", field.BaseIndex, field.DateGrouping.ToString().ToLower(CultureInfo.InvariantCulture));

            if (StartDate.Year < 1900)
            {
                SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", "1900-01-01T00:00:00");
            }
            else
            {
                SetXmlNodeString("d:fieldGroup/d:rangePr/@startDate", StartDate.ToString("s", CultureInfo.InvariantCulture));
                SetXmlNodeString("d:fieldGroup/d:rangePr/@autoStart", "0");
            }

            if (EndDate == DateTime.MaxValue)
            {
                SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", "9999-12-31T00:00:00");
            }
            else
            {
                SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", EndDate.ToString("s", CultureInfo.InvariantCulture));
                SetXmlNodeString("d:fieldGroup/d:rangePr/@autoEnd", "0");
            }

            int items = AddDateGroupItems(group, field.DateGrouping, StartDate, EndDate, interval);

            Grouping = group;
            return group;
        }
        internal ExcelPivotTableFieldNumericGroup SetNumericGroup(int baseIndex, double start, double end, double interval)
        {
            ExcelPivotTableFieldNumericGroup group;
            group = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, TopNode);
            SetXmlNodeBool("d:sharedItems/@containsNumber", true);
            SetXmlNodeBool("d:sharedItems/@containsInteger", true);
            SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);
            SetXmlNodeBool("d:sharedItems/@containsString", false);

            group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"{1}\" endNum=\"{2}\" groupInterval=\"{3}\"/><groupItems /></fieldGroup>",
                baseIndex, start.ToString(CultureInfo.InvariantCulture), end.ToString(CultureInfo.InvariantCulture), interval.ToString(CultureInfo.InvariantCulture));

            int items = AddNumericGroupItems(group, start, end, interval);
            Grouping = group;
            return group;
        }

        private int AddNumericGroupItems(ExcelPivotTableFieldNumericGroup group, double start, double end, double interval)
        {
            if (interval < 0)
            {
                throw (new Exception("The interval must be a positiv"));
            }
            if (start > end)
            {
                throw (new Exception("Then End number must be larger than the Start number"));
            }

            XmlElement groupItemsNode = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
            //First date
            double index = start;
            double nextIndex = start + interval;
            GroupItems.Clear();
            AddGroupItem(groupItemsNode, "<" + start.ToString(CultureInfo.CurrentCulture));

            while (index < end)
            {
                AddGroupItem(groupItemsNode, string.Format("{0}-{1}", index.ToString(CultureInfo.CurrentCulture), nextIndex.ToString(CultureInfo.CurrentCulture)));
                index = nextIndex;
                nextIndex += interval;
                items++;
            }
            AddGroupItem(groupItemsNode, ">" + index.ToString(CultureInfo.CurrentCulture));

            UpdateCacheLookupFromGroupItems();
            return items;
        }
        private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
        {
            XmlElement groupItemsNode = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
            GroupItems.Clear();
            //First date
            AddGroupItem(groupItemsNode, "<" + StartDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));

            switch (GroupBy)
            {
                case eDateGroupBy.Seconds:
                case eDateGroupBy.Minutes:
                    AddTimeSerie(60, groupItemsNode);
                    items += 60;
                    break;
                case eDateGroupBy.Hours:
                    AddTimeSerie(24, groupItemsNode);
                    items += 24;
                    break;
                case eDateGroupBy.Days:
                    if (interval == 1)
                    {
                        DateTime dt = new DateTime(2008, 1, 1); //pick a year with 366 days
                        while (dt.Year == 2008)
                        {
                            AddGroupItem(groupItemsNode, dt.ToString("dd-MMM"));
                            dt = dt.AddDays(1);
                        }
                        items += 366;
                    }
                    else
                    {
                        DateTime dt = StartDate;
                        items = 0;
                        while (dt < EndDate)
                        {
                            AddGroupItem(groupItemsNode, dt.ToString("dd-MMM"));
                            dt = dt.AddDays(interval);
                            items++;
                        }
                    }
                    break;
                case eDateGroupBy.Months:
                    AddGroupItem(groupItemsNode, "jan");
                    AddGroupItem(groupItemsNode, "feb");
                    AddGroupItem(groupItemsNode, "mar");
                    AddGroupItem(groupItemsNode, "apr");
                    AddGroupItem(groupItemsNode, "may");
                    AddGroupItem(groupItemsNode, "jun");
                    AddGroupItem(groupItemsNode, "jul");
                    AddGroupItem(groupItemsNode, "aug");
                    AddGroupItem(groupItemsNode, "sep");
                    AddGroupItem(groupItemsNode, "oct");
                    AddGroupItem(groupItemsNode, "nov");
                    AddGroupItem(groupItemsNode, "dec");
                    items += 12;
                    break;
                case eDateGroupBy.Quarters:
                    AddGroupItem(groupItemsNode, "Qtr1");
                    AddGroupItem(groupItemsNode, "Qtr2");
                    AddGroupItem(groupItemsNode, "Qtr3");
                    AddGroupItem(groupItemsNode, "Qtr4");
                    items += 4;
                    break;
                case eDateGroupBy.Years:
                    if (StartDate.Year >= 1900 && EndDate != DateTime.MaxValue)
                    {
                        for (int year = StartDate.Year; year <= EndDate.Year; year++)
                        {
                            AddGroupItem(groupItemsNode, year.ToString());
                        }
                        items += EndDate.Year - StartDate.Year + 1;
                    }
                    break;
                default:
                    throw (new Exception("unsupported grouping"));
            }

            //Lastdate
            AddGroupItem(groupItemsNode, ">" + EndDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));
            
            UpdateCacheLookupFromGroupItems();

            return items;
        }

        private void UpdateCacheLookupFromGroupItems()
        {
            _cacheLookup = new Dictionary<object, int>();
            for (int i = 0; i < GroupItems.Count; i++)
            {
                _cacheLookup.Add(GroupItems[i], i);
            }
        }

        private void AddTimeSerie(int count, XmlElement groupItems)
        {
            for (int i = 0; i < count; i++)
            {
                AddGroupItem(groupItems, string.Format("{0:00}", i));
            }
        }

        private void AddGroupItem(XmlElement groupItems, string value)
        {
            var s = groupItems.OwnerDocument.CreateElement("s", ExcelPackage.schemaMain);
            s.SetAttribute("v", value);
            groupItems.AppendChild(s);
            GroupItems.Add(value);
        }


        #endregion
        internal void Refresh()
        {
            if (Grouping != null) return;
            var range = _cache.SourceRange;
            var column = range._fromCol + Index;
            var toRow = range._toRow;
            var hs = new HashSet<object>();
            var ws = range.Worksheet;
            //Get unique values.
            for (int row = range._fromRow + 1; row <= toRow; row++)
            {
                var o = ws.GetValue(row, column);
                if (!hs.Contains(o))
                {
                    hs.Add(o);
                }
            }
            //A pivot table cache can reference multiple Pivot tables, so we need to update them all
            foreach (var pt in _cache._pivotTables)
            {
                var existingItems = new HashSet<string>();
                var list = pt.Fields[Index].Items._list;
                var nullItems = 0;
                for (var ix = 0; ix < list.Count; ix++)
                {
                    if (list[ix].Value != null)
                    {
                        if (!hs.Contains(list[ix].Value))
                        {
                            list.RemoveAt(ix);
                            ix--;
                        }
                        else
                        {
                            existingItems.Add(list[ix].Value.ToString());
                        }
                    }
                    else
                    {
                        nullItems++;
                    }
                }
                foreach (var c in hs)
                {
                    if (!existingItems.Contains((c ?? "").ToString()))
                    {
                        list.Insert(list.Count - nullItems, new ExcelPivotTableFieldItem() { Value = c });
                    }
                }
                if (nullItems == 0 && list.Count > 0 && pt.Fields[Index].GetXmlNodeBool("@defaultSubtotal", true) == true)
                {
                    list.Add(new ExcelPivotTableFieldItem() { Type = eItemType.Default, X = -1 });
                }
            }
            SharedItems = hs.ToList();
            if (Slicer != null)
            {
                Slicer.Cache.Data.Items.Refresh();
            }
        }

    }
}