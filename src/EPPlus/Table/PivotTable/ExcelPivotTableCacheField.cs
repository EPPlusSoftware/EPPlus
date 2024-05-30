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
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateAndTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.Utils;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.Serialization;
using System.Xml;
using System.Xml.Linq;

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
            SchemaNodeOrder = ["sharedItems", "fieldGroup", "mpMap"];
            
            if (Grouping == null || Grouping.BaseIndex==Index)
            {
                UpdateCacheLookupFromItems(SharedItems._list, ref _cacheLookup);
            }
            
            if (Grouping != null) 
            {
				UpdateCacheLookupFromItems(GroupItems._list, ref _groupLookup);
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
        internal struct GroupObject<T>
        {
            internal int Index { get; set; }
            internal T MinValue { get; set; }
            internal T MaxValue { get; set; }
        }
        internal Dictionary<object, int> _cacheLookup = null;
		internal Dictionary<object, int> _groupLookup = null;
		internal Dictionary<int, List<int>> _fieldRecordIndex { get; set; }

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
                
                //Grouped fields need to have the max and min values set.
                if (Grouping != null)
                {
                    if (flags == DataTypeFlags.DateTime)
                    {
                        var min = (DateTime)SharedItems.Min();
                        shNode.SetAttribute("minDate", GetDateString(min));
                        var max = (DateTime)SharedItems.Max();
                        shNode.SetAttribute("maxDate", GetDateString(max));
                    }
                    else if ((int)(flags & DataTypeFlags.Number | flags & DataTypeFlags.Int | DataTypeFlags.Float) != 0)
                    {
                        var min = ConvertUtil.GetValueDouble(SharedItems.Min(), true, true);
                        var max = ConvertUtil.GetValueDouble(SharedItems.Max(), true, true);
                        if (!(double.IsNaN(min) || double.IsNaN(max)))
                        {
                            shNode.SetAttribute("minValue", min.ToString(CultureInfo.InvariantCulture));
                            shNode.SetAttribute("maxValue", max.ToString(CultureInfo.InvariantCulture));
                        }
                    }
                }
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
        internal bool IsRowOrColumn
        {
            get
            {
                foreach (var pt in _cache._pivotTables)
                {
                    if (Index < pt.Fields.Count)
                    {
                        var axis = pt.Fields[Index].Axis;
                        if (axis == ePivotFieldAxis.Column ||
                            axis == ePivotFieldAxis.Row)
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
                _formulaTokens = null;
			}
        }
        IList<Token> _formulaTokens = null;
		internal IList<Token> FormulaTokens
        {
            get
            {
                if(_formulaTokens == null && string.IsNullOrEmpty(Formula) == false)
                {
					_formulaTokens = SourceCodeTokenizer.PivotFormula.Tokenize(Formula);
                }
                return _formulaTokens;
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
                if(flags == DataTypeFlags.DateTime)
                {
					shNode.SetAttribute("containsNonDate", "0");
				}
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
                    var t = GetSharedStringText(si, out string dt);
					if (isLongText == false && t.Length > 255) isLongText = true;
                    AppendItem(shNode, dt, t);
				}
            }
            if (isLongText)
            {
                shNode.SetAttribute("longText", "1");
            }
        }

		internal static string GetSharedStringText(object si, out string dt)
		{
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
                        dt = "s";
                        return si.ToString();
					}
					else
					{
                        dt = "n";
                        return ConvertUtil.GetValueForXml(si, false);
					}
				case TypeCode.DateTime:
					dt = "d";
					return GetDateString(((DateTime)si));
				case TypeCode.Boolean:
                    dt = "b";
                    return ConvertUtil.GetValueForXml(si, false);
				case TypeCode.Empty:
                    dt = "m";
                    return null;
				default:
					if (t == typeof(TimeSpan))
					{
                        dt = "d";
                        return GetDateString(new DateTime(((TimeSpan)si).Ticks));
					}
					else if (t == typeof(ExcelErrorValue))
					{
						dt = "e";
						return si.ToString();
					}
					else
					{
						dt = "s";
						return si.ToString();
					}
			}
		}

		private static string GetDateString(DateTime d)
		{
			if (d.Year > 1899)
			{
                return d.ToString("s");
			}
			else
			{
                return d.ToString("HH:mm:ss", CultureInfo.InvariantCulture);
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
                var rangePr = groupNode.SelectSingleNode("d:rangePr", NameSpaceManager);
                if (rangePr != null) 
                {
                    var groupBy = rangePr.Attributes["groupBy"];
                    if (groupBy == null)
                    {
                        Grouping = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, groupNode);
                    }
                    else
                    {
                        DateGrouping = (eDateGroupBy)Enum.Parse(typeof(eDateGroupBy), groupBy.Value, true);
                        Grouping = new ExcelPivotTableFieldDateGroup(NameSpaceManager, groupNode);
                    }
                    var groupItems = groupNode.SelectSingleNode("d:groupItems", NameSpaceManager);
                    if (groupItems != null)
                    {
                        AddItems(GroupItems, groupItems, _groupLookup);
                    }
                }
            }

            var si = GetNode("d:sharedItems");
            if (si != null)
            {
                AddItems(SharedItems, si, _cacheLookup);
            }

        }

        private void AddItems(EPPlusReadOnlyList<Object> items, XmlNode itemsNode, Dictionary<object, int> cacheLookup)
        {
            cacheLookup = new Dictionary<object, int>(new CacheComparer());

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
                
                var key = items[items.Count - 1];
                if (cacheLookup.ContainsKey(key))
                {
                    items._list.Remove(key);
                }
                else
                {
                    cacheLookup.Add(key, items.Count - 1);
                }
            }
        }
        #region Grouping
        internal void SetDateGroup(ExcelPivotTableField field, eDateGroupBy groupBy, DateTime StartDate, DateTime EndDate, int interval, bool firstGroupField)
        {
            if (firstGroupField)
            {
				SetXmlNodeBool("d:sharedItems/@containsDate", true);
				SetXmlNodeBool("d:sharedItems/@containsNonDate", false);
				SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);
			}


			var groupNode = CreateNode("d:fieldGroup"); //Create group topNode
			Grouping = new ExcelPivotTableFieldDateGroup(NameSpaceManager, groupNode);
            
            Grouping.BaseIndex = field.BaseIndex;
			Grouping.TopNode.InnerXml += string.Format("<rangePr groupBy=\"{0}\" /><groupItems />",  groupBy.ToString().ToLower(CultureInfo.InvariantCulture));
			
            if (StartDate == DateTime.MinValue)
			{
				UpdateStartEndValue(out StartDate, out EndDate);
			}

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
                var endDate = EndDate > EndDate.Date ? EndDate.Date.AddDays(1) : EndDate.Date;

				SetXmlNodeString("d:fieldGroup/d:rangePr/@endDate", endDate.ToString("s", CultureInfo.InvariantCulture));
                SetXmlNodeString("d:fieldGroup/d:rangePr/@autoEnd", "0");
            }

			DateGrouping = groupBy;
			
            int items = AddDateGroupItems(Grouping, groupBy, StartDate, EndDate, interval);
        }
        internal void SetNumericGroup(int baseIndex, double start, double end, double interval)
        {
            var groupNode = CreateNode("d:fieldGroup"); //Create group topNode
            var grp = new ExcelPivotTableFieldNumericGroup(NameSpaceManager, groupNode);
			
			SetXmlNodeBool("d:sharedItems/@containsNumber", true);
            SetXmlNodeBool("d:sharedItems/@containsInteger", true);
            SetXmlNodeBool("d:sharedItems/@containsSemiMixedTypes", false);
            SetXmlNodeBool("d:sharedItems/@containsString", false);

			grp.BaseIndex = baseIndex;
			grp.TopNode.InnerXml += string.Format("<rangePr autoStart=\"0\" autoEnd=\"0\" startNum=\"{0}\" endNum=\"{1}\" groupInterval=\"{2}\"/><groupItems />",
                start.ToString(CultureInfo.InvariantCulture), end.ToString(CultureInfo.InvariantCulture), interval.ToString(CultureInfo.InvariantCulture));
            grp.CalculateEndIsDivisibleWithInterval();
            Grouping = grp;
			int items = AddNumericGroupItems(start, end, interval);
        }

        private int AddNumericGroupItems(double start, double end, double interval)
        {
            if (interval < 0)
            {
                throw (new Exception("The interval must be a positiv"));
            }
            if (start > end)
            {
                throw (new Exception("Then End number must be larger than the Start number"));
            }

            XmlElement groupItemsNode = Grouping.TopNode.SelectSingleNode("d:groupItems", Grouping.NameSpaceManager) as XmlElement;
            int items = 2;
            //First date
            double index = start;
            double nextIndex = interval >= 1 ? start + interval-1 : start + interval;
            GroupItems.Clear();
			_groupLookup = new Dictionary<object, int>();

			AddGroupItem(groupItemsNode, "<" + start.ToString(CultureInfo.CurrentCulture));

            while (index < end)
            {
                AddGroupItem(groupItemsNode, string.Format("{0}-{1}", index.ToString(CultureInfo.CurrentCulture), nextIndex.ToString(CultureInfo.CurrentCulture)));
                index = nextIndex;
                nextIndex += (interval >= 1 ? interval - 1 : interval); 
                items++;
            }
            AddGroupItem(groupItemsNode, ">" + index.ToString(CultureInfo.CurrentCulture));

            UpdateCacheLookupFromItems(GroupItems._list, ref _groupLookup);
            return items;
        }
        private int AddDateGroupItems(ExcelPivotTableFieldGroup group, eDateGroupBy GroupBy, DateTime StartDate, DateTime EndDate, int interval)
        {
            XmlElement groupItemsNode = group.TopNode.SelectSingleNode("d:groupItems", group.NameSpaceManager) as XmlElement;
            int items = 2;
			GroupItems.Clear();
            _groupLookup = new Dictionary<object, int>(new CacheComparer());
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
            
            UpdateCacheLookupFromItems(GroupItems._list, ref _groupLookup);

            return items;
        }

        private void UpdateCacheLookupFromItems(List<object> items, ref Dictionary<object, int>  cacheLookup)
        {
            cacheLookup = new Dictionary<object, int>(new CacheComparer());
            for (int i = 0; i < items.Count; i++)
            {
                var key = items[i];
                if (!cacheLookup.ContainsKey(key)) cacheLookup.Add(key, i);
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
            _groupLookup.Add(value, GroupItems.Count);
            GroupItems.Add(value);
        }


        #endregion
        internal void Refresh()
        {
            if (!string.IsNullOrEmpty(Formula)) return;
            if (IsRowColumnOrPage || HasSlicer)
            {
                UpdateSharedItems();
            }
            if(Grouping!=null)
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
                    pt.Fields[Index].UpdateGroupItems(this, true);					
				}
                else
                {
                    pt.Fields[Index].DeleteNode("d:items");
                }
            }
        }

		private void UpdateStartEndValue(out DateTime startDate, out DateTime endDate)
		{
            startDate = DateTime.MaxValue;
            endDate = DateTime.MinValue;
            var ix = Grouping.BaseIndex.Value;
            var fld = _cache.Fields[ix];
            fld.UpdateSharedItems();
            foreach(var item in fld.SharedItems)
            {
                if(item is DateTime dt)
                {
					if(startDate > dt) startDate = dt;
                    if(endDate < dt) endDate = dt;
				}
            }
		}

		private void UpdateSharedItems()
		{
			var range = _cache.SourceRange;
			if (range == null) return;
			var column = range._fromCol + Index;
			var siHs = new HashSet<object>(new InvariantObjectComparer());
			var ws = range.Worksheet;
			int toRow = _cache.GetMaxRow();

			//Get unique values.
			for (int row = range._fromRow + 1; row <= toRow; row++)
			{
				AddSharedItemToHashSet(siHs, ws.GetValue(row, column));
			}

            if (Grouping == null)
            {
                UpdatePivotItemsFromSharedItems(siHs);
            }
            SharedItems._list = siHs.ToList();
			UpdateCacheLookupFromItems(SharedItems._list, ref _cacheLookup);
			if (HasSlicer)
			{
				UpdateSlicers();
			}
		}

        private void UpdatePivotItemsFromSharedItems(HashSet<object> siHs)
        {
            //A pivot table cache can reference multiple Pivot tables, so we need to update them all
            foreach (var pt in _cache._pivotTables)
            {
                var ptField = pt.Fields[Index];
                if (ptField.ShouldHaveItems == false) continue;
                var existingItems = new HashSet<object>(new InvariantObjectComparer());
                var list = ptField.Items._list;

                for (var ix = 0; ix < list.Count; ix++)
                {
                    var v = list[ix].Value ?? ExcelPivotTable.PivotNullValue;
                    if (!siHs.Contains(v) || existingItems.Contains(v))
                    {
                        list.RemoveAt(ix);
                        ix--;
                    }
                    else
                    {
                        existingItems.Add(v);
                    }
                }
                var hasSubTotalSubt = list.Count > 0 && list[list.Count - 1].Type == eItemType.Default ? 1 : 0;
                foreach (var c in siHs)
                {
                    if (!existingItems.Contains(c))
                    {
                        list.Insert(list.Count - hasSubTotalSubt, new ExcelPivotTableFieldItem() { Value = c });
                    }
                }

                if (list.Count > 0)
                {
                    UpdateSubTotalItems(list, ptField.SubTotalFunctions);
                }
            }
        }

        internal void UpdateSubTotalItems(List<ExcelPivotTableFieldItem> list, eSubTotalFunctions functions)
        {
            while(list.Count > 0 && list[list.Count-1].Type!=eItemType.Data) { list.RemoveAt(list.Count - 1); }
            if (functions == eSubTotalFunctions.None) return;
            foreach (var t in GetItemTypeFromFunctionList(functions))
            {
                list.Add(new ExcelPivotTableFieldItem() { Type = t, X = -1 }); 
            }
        }

        private List<eItemType> GetItemTypeFromFunctionList(eSubTotalFunctions subTotalFunctions)
        {
            var l = new List<eItemType>();
            foreach(eSubTotalFunctions t in Enum.GetValues(typeof(eSubTotalFunctions)))
            {
                if ((t & subTotalFunctions) != 0)
                {
                    switch (t)
                    {
                        case eSubTotalFunctions.Sum:
                            l.Add(eItemType.Sum);
                            break;
                        case eSubTotalFunctions.Min:
                            l.Add(eItemType.Min);
                            break;
                        case eSubTotalFunctions.Max:
                            l.Add(eItemType.Max);
                            break;
                        case eSubTotalFunctions.Avg:
                            l.Add(eItemType.Avg);
                            break;
                        case eSubTotalFunctions.Count:
                            l.Add(eItemType.Count);
                            break;
                        case eSubTotalFunctions.CountA:
                            l.Add(eItemType.CountA);
                            break;
                        case eSubTotalFunctions.Product:
                            l.Add(eItemType.Product);
                            break;
                        case eSubTotalFunctions.StdDev:
                            l.Add(eItemType.StdDev);
                            break;
                        case eSubTotalFunctions.StdDevP:
                            l.Add(eItemType.StdDevP);
                            break;
                        case eSubTotalFunctions.Var:
                            l.Add(eItemType.Var);
                            break;
                        case eSubTotalFunctions.VarP:
                            l.Add(eItemType.VarP);
                            break;
                        default:
                            l.Add(eItemType.Default);
                            break;
                    }
                }
            }
            return l;
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

		internal Dictionary<object, int> GetCacheLookup()
		{
			if(Grouping == null)
            {
                return _cacheLookup;
            }
            else
            {
                return _groupLookup;
            }
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
            if (x == null || x.Equals(ExcelPivotTable.PivotNullValue)) return ExcelPivotTable.PivotNullValue;

			if (x is string sx)
            {
				return sx.ToLower();
			}
            else if (x is char cx)
            {
                return char.ToLower(cx).ToString();
            }
            if(ConvertUtil.IsExcelNumeric(x))
            {
                return ConvertUtil.GetValueDouble(x).ToString(CultureInfo.InvariantCulture);
            }
            return x.ToString().ToLower();
        }

        public int GetHashCode(object obj)
        {
            return GetCaseInsensitiveValue(obj).GetHashCode();
        }
    }
}