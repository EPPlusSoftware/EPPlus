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
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Globalization;
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
        private int _index;
        internal ExcelPivotTableCacheField(XmlNamespaceManager nsm, XmlNode topNode,  PivotTableCacheInternal cache, int index) : base(nsm, topNode)
        {
            _cache = cache;
            _index = index;
            SetCacheFieldNode();
        }
        public string Name
        {
            get;
            internal set;
        }
        public List<object> SharedItems
        {
            get;
            set;
        } = new List<object>();
        internal Dictionary<object, int> _cacheLookup=null;

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
                if ((flags & DataTypeFlags.String) != DataTypeFlags.String)
                {
                    shNode.SetAttribute("containsSemiMixedTypes", "0");
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
                    var axis = pt.Fields[_index].Axis;
                    if (axis == ePivotFieldAxis.Column || 
                        axis == ePivotFieldAxis.Row ||
                        axis == ePivotFieldAxis.Page)
                    {
                        return true;
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
                            AppendItem(shNode, "d", ConvertUtil.GetValueForXml(si, false));
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
            }

            var si = GetNode("d:sharedItems");
            if (si != null)
            {
                foreach (XmlElement c in si.ChildNodes)
                {
                    if (c.LocalName == "s")
                    {
                        SharedItems.Add(c.Attributes["v"].Value);
                    }
                    else if (c.LocalName == "d")
                    {
                        if (ConvertUtil.TryParseDateString(c.Attributes["v"].Value, out DateTime d))
                        {
                            SharedItems.Add(d);
                        }
                        else
                        {
                            SharedItems.Add(c.Attributes["v"].Value);
                        }
                    }
                    else if (c.LocalName == "n")
                    {
                        if (ConvertUtil.TryParseNumericString(c.Attributes["v"].Value, out double num))
                        {
                            SharedItems.Add(num);
                        }
                        else
                        {
                            SharedItems.Add(c.Attributes["v"].Value);
                        }
                    }
                    else if (c.LocalName == "b")
                    {
                        if (ConvertUtil.TryParseBooleanString(c.Attributes["v"].Value, out bool b))
                        {
                            SharedItems.Add(b);
                        }
                        else
                        {
                            SharedItems.Add(c.Attributes["v"].Value);
                        }
                    }
                    else if (c.LocalName == "e")
                    {
                        if (ExcelErrorValue.Values.StringIsErrorValue(c.Attributes["v"].Value))
                        {
                            SharedItems.Add(ExcelErrorValue.Parse(c.Attributes["v"].Value));
                        }
                        else
                        {
                            SharedItems.Add(c.Attributes["v"].Value);
                        }
                    }
                    else
                    {
                        SharedItems.Add(null);
                    }
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
            AddFieldItems(items);

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

            XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
            //First date
            double index = start;
            double nextIndex = start + interval;
            AddGroupItem(groupItems, "<" + start.ToString(CultureInfo.InvariantCulture));

            while (index < end)
            {
                AddGroupItem(groupItems, string.Format("{0}-{1}", index.ToString(CultureInfo.InvariantCulture), nextIndex.ToString(CultureInfo.InvariantCulture)));
                index = nextIndex;
                nextIndex += interval;
                items++;
            }
            AddGroupItem(groupItems, ">" + nextIndex.ToString(CultureInfo.InvariantCulture));
            return items;
        }

        private void AddFieldItems(int items)
        {
            XmlElement prevNode = null;
            XmlElement itemsNode = TopNode.SelectSingleNode("d:items", NameSpaceManager) as XmlElement;
            for (int x = 0; x < items; x++)
            {
                var itemNode = itemsNode.OwnerDocument.CreateElement("item", ExcelPackage.schemaMain);
                itemNode.SetAttribute("x", x.ToString());
                if (prevNode == null)
                {
                    itemsNode.PrependChild(itemNode);
                }
                else
                {
                    itemsNode.InsertAfter(itemNode, prevNode);
                }
                prevNode = itemNode;
            }
            itemsNode.SetAttribute("count", (items + 1).ToString());
        }

        private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
        {
            XmlElement groupItems = group.TopNode.SelectSingleNode("d:fieldGroup/d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
            //First date
            AddGroupItem(groupItems, "<" + StartDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));

            switch (GroupBy)
            {
                case eDateGroupBy.Seconds:
                case eDateGroupBy.Minutes:
                    AddTimeSerie(60, groupItems);
                    items += 60;
                    break;
                case eDateGroupBy.Hours:
                    AddTimeSerie(24, groupItems);
                    items += 24;
                    break;
                case eDateGroupBy.Days:
                    if (interval == 1)
                    {
                        DateTime dt = new DateTime(2008, 1, 1); //pick a year with 366 days
                        while (dt.Year == 2008)
                        {
                            AddGroupItem(groupItems, dt.ToString("dd-MMM"));
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
                            AddGroupItem(groupItems, dt.ToString("dd-MMM"));
                            dt = dt.AddDays(interval);
                            items++;
                        }
                    }
                    break;
                case eDateGroupBy.Months:
                    AddGroupItem(groupItems, "jan");
                    AddGroupItem(groupItems, "feb");
                    AddGroupItem(groupItems, "mar");
                    AddGroupItem(groupItems, "apr");
                    AddGroupItem(groupItems, "may");
                    AddGroupItem(groupItems, "jun");
                    AddGroupItem(groupItems, "jul");
                    AddGroupItem(groupItems, "aug");
                    AddGroupItem(groupItems, "sep");
                    AddGroupItem(groupItems, "oct");
                    AddGroupItem(groupItems, "nov");
                    AddGroupItem(groupItems, "dec");
                    items += 12;
                    break;
                case eDateGroupBy.Quarters:
                    AddGroupItem(groupItems, "Qtr1");
                    AddGroupItem(groupItems, "Qtr2");
                    AddGroupItem(groupItems, "Qtr3");
                    AddGroupItem(groupItems, "Qtr4");
                    items += 4;
                    break;
                case eDateGroupBy.Years:
                    if (StartDate.Year >= 1900 && EndDate != DateTime.MaxValue)
                    {
                        for (int year = StartDate.Year; year <= EndDate.Year; year++)
                        {
                            AddGroupItem(groupItems, year.ToString());
                        }
                        items += EndDate.Year - StartDate.Year + 1;
                    }
                    break;
                default:
                    throw (new Exception("unsupported grouping"));
            }

            //Lastdate
            AddGroupItem(groupItems, ">" + EndDate.ToString("s", CultureInfo.InvariantCulture).Substring(0, 10));
            return items;
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
        }

        #endregion
    }
}