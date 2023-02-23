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
using OfficeOpenXml.Constants;
using OfficeOpenXml.Core;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Slicer;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// A pivot tables cache field
    /// </summary>
    public class ExcelPivotTableCacheField : XmlHelper
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
            Error = 0x30,
            Float = 0x40,
        }
        internal PivotTableCacheInternal _cache;
        internal ExcelPivotTableCacheField(XmlNamespaceManager nsm, XmlNode topNode, PivotTableCacheInternal cache, int index) : base(nsm, topNode)
        {
            _cache = cache;
            Index = index;
            SetCacheFieldNode();
            if (NumFmtId.HasValue)
            {
                var styles = cache._wb.Styles;
                var ix = styles.NumberFormats.FindIndexById(NumFmtId.Value.ToString(CultureInfo.InvariantCulture));
                if (ix >= 0)
                {
                    Format = styles.NumberFormats[ix].Format;
                }
            }
        }
        /// <summary>
        /// The index in the collection of the pivot field
        /// </summary>
        public int Index { get; set; }
        /// <summary>
        /// The name for the field
        /// </summary>
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
        public EPPlusReadOnlyList<object> SharedItems
        {
            get;
        } = new EPPlusReadOnlyList<object>();
        /// <summary>
        /// A list of group items, if the field has grouping.
        /// <seealso cref="Grouping"/>
        /// </summary>
        public EPPlusReadOnlyList<object> GroupItems
        {
            get;
            set;
        } = new EPPlusReadOnlyList<object>();
        internal Dictionary<object, int> _cacheLookup = null;
        /// <summary>
        /// The type of date grouping
        /// </summary>
        public eDateGroupBy DateGrouping { get; private set; }
        /// <summary>
        /// Grouping proprerties, if the field has grouping
        /// </summary>
        public ExcelPivotTableFieldGroup Grouping { get; set; }
        /// <summary>
        /// The number format for the field
        /// </summary>
        public string Format { get; set; }
        internal int? NumFmtId
        {
            get
            {
                return GetXmlNodeIntNull("@numFmtId");
            }
            set
            {
                SetXmlNodeInt("@numFmtId", value);
            }
        }
        internal void WriteSharedItems(XmlElement fieldNode, XmlNamespaceManager nsm)
        {
            var shNode = (XmlElement)fieldNode.SelectSingleNode("d:sharedItems", nsm);
            shNode.RemoveAll();

            var flags = GetFlags();

            _cacheLookup = new Dictionary<object, int>(new CacheComparer());
            if (IsRowColumnOrPage || HasSlicer)
            {
                AppendSharedItems(shNode);
            }
            var noTypes = GetNoOfTypes(flags);
            if (noTypes > 1 && 
                flags != (DataTypeFlags.Int | DataTypeFlags.Number) &&
                flags != (DataTypeFlags.Float | DataTypeFlags.Number) &&
                flags != (DataTypeFlags.Int | DataTypeFlags.Float | DataTypeFlags.Number) &&
                flags != (DataTypeFlags.Int | DataTypeFlags.Number | DataTypeFlags.Empty) &&
                flags != (DataTypeFlags.Float | DataTypeFlags.Number | DataTypeFlags.Empty) &&
                flags != (DataTypeFlags.Int | DataTypeFlags.Float | DataTypeFlags.Number | DataTypeFlags.Empty) &&
                SharedItems.Count > 1)
            {
                if ((flags & DataTypeFlags.String) == DataTypeFlags.String ||
                    (flags & DataTypeFlags.String) == DataTypeFlags.Empty)
                {
                    shNode.SetAttribute("containsMixedTypes", "1");
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
                    (flags & DataTypeFlags.Empty) != DataTypeFlags.Empty &&
                    (flags & DataTypeFlags.Boolean) != DataTypeFlags.Boolean)
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
                foreach (var pt in _cache._pivotTables)
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
        /// <summary>
        /// The formula for cache field.
        /// The formula for the calculated field. 
        /// Note: In formulas you create for calculated fields or calculated items, you can use operators and expressions as you do in other worksheet formulas. 
        /// You can use constants and refer to data from the pivot table, but you cannot use cell references or defined names.You cannot use worksheet functions that require cell references or defined names as arguments, and you cannot use array functions.
        /// <seealso cref="ExcelPivotTableFieldCollection.AddCalculatedField(string, string)"/>
        /// </summary>
        public string Formula
        {
            get
            {
                return GetXmlNodeString("@formula");
            }
            set
            {
                if(DatabaseField)
                {
                    throw new InvalidOperationException("Can't set a formula to a database field");
                }
                if (string.IsNullOrEmpty(value) || value.Trim() == "")
                {
                    throw (new ArgumentException("The formula can't be blank", "formula"));
                }
                SetXmlNodeString("@formula", value);
            }
        }
        internal bool DatabaseField
        {
            get
            {
                return GetXmlNodeBool("@databaseField", true);
            }
            set
            {
                SetXmlNodeBool("@databaseField", value, true);
            }
        }
        internal bool HasSlicer
        {
            get
            {
                foreach (var pt in _cache._pivotTables)
                {
                    if (pt.Fields.Count>Index && pt.Fields[Index].Slicer != null)
                    {
                        return true;
                    }
                }
                return false;
            }
        }
        internal void UpdateSlicers()
        {
            foreach (var pt in _cache._pivotTables)
            {
                var s = pt.Fields[Index].Slicer;
                if (s != null)
                {
                    s.Cache.Data.Items.RefreshMe();
                }
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
            if ((flags & DataTypeFlags.Int) == DataTypeFlags.Int &&
                (flags & DataTypeFlags.Float) != DataTypeFlags.Float)
            {
                shNode.SetAttribute("containsInteger", "1");
            }
            if ((flags & DataTypeFlags.Empty) == DataTypeFlags.Empty)
            {
                shNode.SetAttribute("containsBlank", "1");
            }
            if((flags & DataTypeFlags.String) != DataTypeFlags.String &&
               (flags & DataTypeFlags.Boolean) != DataTypeFlags.Boolean &&
               (flags & DataTypeFlags.Error) != DataTypeFlags.Error)
            {
                shNode.SetAttribute("containsString", "0");
            }
        }
        private int GetNoOfTypes(DataTypeFlags flags)
        {
            int types = 0;
            foreach (DataTypeFlags v in Enum.GetValues(typeof(DataTypeFlags)))
            {
                if (v!=DataTypeFlags.Empty && (flags & v)==v)
                {
                    types++;
                }
            }
            return types;
        }
        private void AppendSharedItems(XmlElement shNode)
        {
            int index = 0;
            bool isLongText = false;
            foreach (var si in SharedItems)
            {
                if (si == null || si.Equals(ExcelPivotTable.PivotNullValue))
                {
                    _cacheLookup.Add(ExcelPivotTable.PivotNullValue, index++);
                    AppendItem(shNode, "m", null);
                }
                else
                {
                    _cacheLookup.Add(si, index++);
                    var t = si.GetType();
                    var tc = Type.GetTypeCode(t);

                    switch (tc)
                    {
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
                            if (t.IsEnum)
                            {
                                AppendItem(shNode, "s", si.ToString());
                            }
                            else
                            {
                                AppendItem(shNode, "n", ConvertUtil.GetValueForXml(si, false));
                            }
                            break;
                        case TypeCode.DateTime:
                            var d = ((DateTime)si);
                            if (d.Year > 1899)
                            {
                                AppendItem(shNode, "d", d.ToString("s"));
                            }
                            else
                            {
                                AppendItem(shNode, "d", d.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                            }
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
                                d = new DateTime(((TimeSpan)si).Ticks);
                                if (d.Year > 1899)
                                {
                                    AppendItem(shNode, "d", d.ToString("s"));
                                }
                                else
                                {
                                    AppendItem(shNode, "d", d.ToString("HH:mm:ss", CultureInfo.InvariantCulture));
                                }
                            }
                            else if (t == typeof(ExcelErrorValue))
                            {
                                AppendItem(shNode, "e", si.ToString());
                            }
                            else
                            {
                                var s = si.ToString();
                                AppendItem(shNode, "s", s);
                                if (s.Length > 255 && isLongText == false) isLongText = true;
                            }
                            break;
                    }
                }
            }
            if (isLongText)
            {
                shNode.SetAttribute("longText", "1");
            }
        }

        private DataTypeFlags GetFlags()
        {
            DataTypeFlags flags = 0;
            foreach (var si in SharedItems)
            {
                if (si == null || si.Equals(ExcelPivotTable.PivotNullValue))
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
                            if(t.IsEnum)
                            {
                                flags |= DataTypeFlags.String;
                            }
                            else
                            {
                                flags |= (DataTypeFlags.Number | DataTypeFlags.Int);
                            }
                            break;
                        case TypeCode.Decimal:
                        case TypeCode.Double:
                        case TypeCode.Single:
                            flags |= (DataTypeFlags.Number);
                            if ((flags & DataTypeFlags.Int) != DataTypeFlags.Int && (Convert.ToDouble(si) % 1 == 0))
                            {
                                flags |= DataTypeFlags.Int;
                            }
                            else if ((flags & DataTypeFlags.Float) != DataTypeFlags.Float && (Convert.ToDouble(si) % 1 != 0))
                            {
                                flags |= DataTypeFlags.Float;
                            }
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
                var groupItems = groupNode.SelectSingleNode("d:groupItems", NameSpaceManager);
                if (groupItems != null)
                {
                    AddItems(GroupItems, groupItems, true);
                }
            }

            var si = GetNode("d:sharedItems");
            if (si != null)
            {
                AddItems(SharedItems, si, groupNode==null);
            }

        }

        private void AddItems(EPPlusReadOnlyList<Object> items, XmlNode itemsNode, bool updateCacheLookup)
        {
            if (updateCacheLookup)
            {
                _cacheLookup = new Dictionary<object, int>(new CacheComparer());
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
                    items.Add(ExcelPivotTable.PivotNullValue);
                }
                if(updateCacheLookup)
                {
                    var key = items[items.Count - 1];
                    if (_cacheLookup.ContainsKey(key))
                    {
                        items._list.Remove(key);
                    }
                    else
                    {
                        _cacheLookup.Add(key, items.Count - 1);
                    }
                }
            }
        }
        #region Grouping
        internal ExcelPivotTableFieldDateGroup SetDateGroup(ExcelPivotTableField field, eDateGroupBy groupBy, DateTime StartDate, DateTime EndDate, int interval)
        {
            ExcelPivotTableFieldDateGroup group;
            group = new ExcelPivotTableFieldDateGroup(NameSpaceManager, TopNode);
            SetXmlNodeBool("d:sharedItems/@containsDate", true);
            SetXmlNodeBool("d:sharedItems/@containsNonDate", false);
            SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);

            group.TopNode.InnerXml += string.Format("<fieldGroup base=\"{0}\"><rangePr groupBy=\"{1}\" /><groupItems /></fieldGroup>", field.BaseIndex, groupBy.ToString().ToLower(CultureInfo.InvariantCulture));

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

            int items = AddDateGroupItems(group, groupBy, StartDate, EndDate, interval);

            Grouping = group;
            DateGrouping = groupBy;
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

            UpdateCacheLookupFromItems(GroupItems._list);
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
            
            UpdateCacheLookupFromItems(GroupItems._list);

            return items;
        }

        private void UpdateCacheLookupFromItems(List<object> items)
        {
            _cacheLookup = new Dictionary<object, int>(new CacheComparer());
            for (int i = 0; i < items.Count; i++)
            {
                var key = items[i];
                if (!_cacheLookup.ContainsKey(key)) _cacheLookup.Add(key, i);
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
            if (!string.IsNullOrEmpty(Formula)) return;
            if (Grouping == null)
            {
                UpdateSharedItems();
            }
            else
            {
                UpdateGroupItems();
            }
        }

        private void UpdateGroupItems()
        {
            foreach (var pt in _cache._pivotTables)
            {                
                if ((pt.Fields[Index].IsRowField ||
                     pt.Fields[Index].IsColumnField ||
                     pt.Fields[Index].IsPageField || pt.Fields[Index].Cache.HasSlicer) )
                {
                    if (pt.Fields[Index].Items.Count == 0)
                    {
                        pt.Fields[Index].UpdateGroupItems(this, true);
                    }
                }
                else
                {
                    pt.Fields[Index].DeleteNode("d:items");
                }
            }
        }

        private void UpdateSharedItems()
        {
            var range = _cache.SourceRange;
            if (range == null) return;
            var column = range._fromCol + Index;
            var hs = new HashSet<object>(new InvariantObjectComparer());
            var ws = range.Worksheet;
            var dimensionToRow = ws.Dimension?._toRow ?? range._fromRow + 1;
            var toRow = range._toRow < dimensionToRow ? range._toRow : dimensionToRow;

            //Get unique values.
            for (int row = range._fromRow + 1; row <= toRow; row++)
            {
                AddSharedItemToHashSet(hs, ws.GetValue(row, column));
            }
            //A pivot table cache can reference multiple Pivot tables, so we need to update them all
            foreach (var pt in _cache._pivotTables)
            {
                var existingItems = new HashSet<object>();
                var list = pt.Fields[Index].Items._list;
                
                for (var ix = 0; ix < list.Count; ix++)
                {
                    var v = list[ix].Value ?? ExcelPivotTable.PivotNullValue;
                    if (!hs.Contains(v) || existingItems.Contains(v))
                    {
                        list.RemoveAt(ix);
                        ix--;
                    }
                    else
                    {
                        existingItems.Add(v);
                    }
                }
                var hasSubTotalSubt=list.Count > 0 && list[list.Count-1].Type==eItemType.Default ? 1 : 0;
                foreach (var c in hs)
                {
                    if (!existingItems.Contains(c))
                    {
                        list.Insert(list.Count - hasSubTotalSubt, new ExcelPivotTableFieldItem() { Value = c });
                    }
                }

                if (list.Count > 0 && list[list.Count - 1].Type != eItemType.Default && pt.Fields[Index].GetXmlNodeBool("@defaultSubtotal", true) == true)
                {
                    list.Add(new ExcelPivotTableFieldItem() { Type = eItemType.Default, X = -1 });
                }
            }
            SharedItems._list = hs.ToList();
            UpdateCacheLookupFromItems(SharedItems._list);
            if (HasSlicer)
            {
                UpdateSlicers();
            }
        }
        internal static object AddSharedItemToHashSet(HashSet<object> hs, object o)
        {
            if (o == null)
            {
                o = ExcelPivotTable.PivotNullValue;
            }
            else
            {
                var t = o.GetType();
                if (t == typeof(TimeSpan))
                {
                    var ticks = ((TimeSpan)o).Ticks + (TimeSpan.TicksPerSecond) / 2;
                    o = new DateTime(ticks - (ticks % TimeSpan.TicksPerSecond));
                }
                if (t == typeof(DateTime))
                {
                    var ticks = ((DateTime)o).Ticks;
                    if ((ticks % TimeSpan.TicksPerSecond) != 0)
                    {
                        ticks += TimeSpan.TicksPerSecond / 2;
                        o = new DateTime(ticks - (ticks % TimeSpan.TicksPerSecond));
                    }
                }
            }
            if (!hs.Contains(o))
            {
                hs.Add(o);
            }

            return o;
        }

    }

    internal class CacheComparer : IEqualityComparer<object>
    {
        public new bool Equals(object x, object y)
        {
            x = GetCaseInsensitiveValue(x);
            y = GetCaseInsensitiveValue(y);

            return x.Equals(y);
        }

        private static object GetCaseInsensitiveValue(object x)
        {
            if (x is string sx)
            {
                x = sx.ToLower();
            }
            else if (x is char cx)
            {
                x = char.ToLower(cx);
            }

            return x;
        }

        public int GetHashCode(object obj)
        {
            return GetCaseInsensitiveValue(obj).GetHashCode();
        }
    }
}