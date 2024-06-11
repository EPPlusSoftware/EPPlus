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
using System.Text;
using System.Xml;
using System.Text.RegularExpressions;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using OfficeOpenXml.Packaging;
using System.Linq;
using OfficeOpenXml.Constants;
using OfficeOpenXml.Filter;
using OfficeOpenXml.Packaging.Ionic;
using OfficeOpenXml.FormulaParsing.Excel.Functions.MathFunctions;
using OfficeOpenXml.Style.Dxf;
using System.IO;
using System.Globalization;
using OfficeOpenXml.Table.PivotTable.Filter;
using OfficeOpenXml.Table.PivotTable.Calculation;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Finance;
using OfficeOpenXml.Table.PivotTable.Calculation.ShowDataAs;

namespace OfficeOpenXml.Table.PivotTable
{
    /// <summary>
    /// Represents a null value in a pivot table caches shared items list.
    /// </summary>
    public struct PivotNull : IEqualityComparer<PivotNull>
	{
		public bool Equals(PivotNull x, PivotNull y)
		{
            return true;
		}

		public int GetHashCode(PivotNull obj)
		{
			return 0;
		}

		/// <summary>
		/// Return the string representation of the pivot null value
		/// </summary>
		/// <returns>An empty string</returns>
		public override string ToString()
		{
			return "";
		}
	}
    /// <summary>
    /// An Excel Pivottable
    /// </summary>
    public class ExcelPivotTable : XmlHelper
    {
        /// <summary>
        /// Represents a null value in a pivot table caches shared items list.
        /// </summary>
        public static PivotNull PivotNullValue = new PivotNull();
        internal ExcelPivotTable(ZipPackageRelationship rel, ExcelWorksheet sheet) :
            base(sheet.NameSpaceManager)
        {
            WorkSheet = sheet;
            PivotTableUri = UriHelper.ResolvePartUri(rel.SourceUri, rel.TargetUri);
            Relationship = rel;
            var pck = sheet._package.ZipPackage;
            Part = pck.GetPart(PivotTableUri);

            PivotTableXml = new XmlDocument();
            LoadXmlSafe(PivotTableXml, Part.GetStream());
            TopNode = PivotTableXml.DocumentElement;
            Init();
            Address = new ExcelAddressBase(GetXmlNodeString("d:location/@ref"));

            CacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this);
            LoadFields();

            var pos = 0;
            //Add row fields.
            foreach (XmlElement rowElem in TopNode.SelectNodes("d:rowFields/d:field", NameSpaceManager))
            {
                if (int.TryParse(rowElem.GetAttribute("x"), out int x) && x >= 0)
                {
                    RowFields.AddInternal(Fields[x]);
                }
                else
                {
                    if (x == -2)
                    {
                        ValuesFieldPosition = pos;
                    }
                    rowElem.ParentNode.RemoveChild(rowElem);
                }
                pos++;
            }

            pos = 0;
            ////Add column fields.
            foreach (XmlElement colElem in TopNode.SelectNodes("d:colFields/d:field", NameSpaceManager))
            {
                if (int.TryParse(colElem.GetAttribute("x"), out int x) && x >= 0)
                {
                    ColumnFields.AddInternal(Fields[x]);
                }
                else
                {
                    if (x == -2)
                    {
                        ValuesFieldPosition = pos;
                    }
                    colElem.ParentNode.RemoveChild(colElem);
                }
                pos++;
            }

            //Add Page elements
            //int index = 0;
            foreach (XmlElement pageElem in TopNode.SelectNodes("d:pageFields/d:pageField", NameSpaceManager))
            {
                if (int.TryParse(pageElem.GetAttribute("fld"), out int fld) && fld >= 0)
                {
                    var field = Fields[fld];
                    field._pageFieldSettings = new ExcelPivotTablePageFieldSettings(NameSpaceManager, pageElem, field, fld);
                    PageFields.AddInternal(field);
                }
            }

            //Add data elements
            //index = 0;
            foreach (XmlElement dataElem in TopNode.SelectNodes("d:dataFields/d:dataField", NameSpaceManager))
            {
                if (int.TryParse(dataElem.GetAttribute("fld"), out int fld) && fld >= 0)
                {
                    var field = Fields[fld];
                    var dataField = new ExcelPivotTableDataField(NameSpaceManager, dataElem, field);
                    DataFields.AddInternal(dataField);
                }
            }

            Styles = new ExcelPivotTableAreaStyleCollection(this);
        }
        /// <summary>
        /// Add a new pivottable
        /// </summary>
        /// <param name="sheet">The worksheet</param>
        /// <param name="address">the address of the pivottable</param>
        /// <param name="pivotTableCache">The pivot table cache</param>
        /// <param name="name"></param>
        /// <param name="tblId"></param>
        internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddressBase address, PivotTableCacheInternal pivotTableCache, string name, int tblId) :
        base(sheet.NameSpaceManager)
        {
            CreatePivotTable(sheet, address, pivotTableCache.Fields.Count, name, tblId);

            CacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, pivotTableCache);
            CacheId = pivotTableCache.ExtLstCacheId;

            LoadFields();
            Styles = new ExcelPivotTableAreaStyleCollection(this);
        }
        /// <summary>
        /// Add a new pivottable
        /// </summary>
        /// <param name="sheet">The worksheet</param>
        /// <param name="address">the address of the pivottable</param>
        /// <param name="sourceAddress">The address of the Source data</param>
        /// <param name="name"></param>
        /// <param name="tblId"></param>
        internal ExcelPivotTable(ExcelWorksheet sheet, ExcelAddressBase address, ExcelRangeBase sourceAddress, string name, int tblId) :
        base(sheet.NameSpaceManager)
        {
            CreatePivotTable(sheet, address, sourceAddress._toCol - sourceAddress._fromCol + 1, name, tblId);

            CacheDefinition = new ExcelPivotCacheDefinition(sheet.NameSpaceManager, this, sourceAddress);
            CacheId = CacheDefinition._cacheReference.ExtLstCacheId;

            LoadFields();
            Styles = new ExcelPivotTableAreaStyleCollection(this);
        }

        private void CreatePivotTable(ExcelWorksheet sheet, ExcelAddressBase address, int fields, string name, int tblId)
        {
            WorkSheet = sheet;
            Address = address;
            var pck = sheet._package.ZipPackage;

            PivotTableXml = new XmlDocument();
            LoadXmlSafe(PivotTableXml, GetStartXml(name, address, fields), Encoding.UTF8);
            TopNode = PivotTableXml.DocumentElement;
            PivotTableUri = GetNewUri(pck, "/xl/pivotTables/pivotTable{0}.xml", ref tblId);
            Init();

            Part = pck.CreatePart(PivotTableUri, ContentTypes.contentTypePivotTable);
            PivotTableXml.Save(Part.GetStream());

            //Worksheet-Pivottable relationship
            Relationship = sheet.Part.CreateRelationship(UriHelper.ResolvePartUri(sheet.WorksheetUri, PivotTableUri), Packaging.TargetMode.Internal, ExcelPackage.schemaRelationships + "/pivotTable");

            using (var r = sheet.Cells[address.Address])
            {
                r.Clear();
            }
        }

        private void Init()
        {
            SchemaNodeOrder = new string[] { "location", "pivotFields", "rowFields", "rowItems", "colFields", "colItems", "pageFields", "dataFields", "formats", "conditionalFormats", "chartFormats", "pivotHierarchies", "pivotTableStyleInfo", "filters", "rowHierarchiesUsage", "colHierarchiesUsage", "extLst" };
        }
        private void LoadFields()
        {
            int index = 0;
            var pivotFieldNode = TopNode.SelectSingleNode("d:pivotFields", NameSpaceManager);
            //Add fields.            
            foreach (XmlElement fieldElem in pivotFieldNode.SelectNodes("d:pivotField", NameSpaceManager))
            {
                var fld = new ExcelPivotTableField(NameSpaceManager, fieldElem, this, index, index);
                fld.Cache = CacheDefinition._cacheReference.Fields[index++];
                fld.LoadItems();
                Fields.AddInternal(fld);
            }

        }
        private string GetStartXml(string name, ExcelAddressBase address, int fields)
        {
            string xml = string.Format("<pivotTableDefinition xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" name=\"{0}\" dataOnRows=\"1\" applyNumberFormats=\"0\" applyBorderFormats=\"0\" applyFontFormats=\"0\" applyPatternFormats=\"0\" applyAlignmentFormats=\"0\" applyWidthHeightFormats=\"1\" dataCaption=\"Data\"  createdVersion=\"6\" updatedVersion=\"6\" showMemberPropertyTips=\"0\" useAutoFormatting=\"1\" itemPrintTitles=\"1\" indent=\"0\" compact=\"0\" compactData=\"0\" gridDropZones=\"1\">",
                ConvertUtil.ExcelEscapeString(name));

            xml += string.Format("<location ref=\"{0}\" firstHeaderRow=\"1\" firstDataRow=\"1\" firstDataCol=\"1\" /> ", address.FirstAddress);
            xml += string.Format("<pivotFields count=\"{0}\">", fields);
            for (int col = 0; col < fields; col++)
            {
                xml += "<pivotField showAll=\"0\" />"; //compact=\"0\" outline=\"0\" subtotalTop=\"0\" includeNewItemsInFilter=\"1\"     
            }

            xml += "</pivotFields>";
            xml += "<pivotTableStyleInfo name=\"PivotStyleMedium9\" showRowHeaders=\"1\" showColHeaders=\"1\" showRowStripes=\"0\" showColStripes=\"0\" showLastColumn=\"1\" />";
            xml += $"<extLst><ext xmlns:xpdl=\"http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout\" uri=\"{ExtLstUris.PivotTableDefinition16Uri}\"><xpdl:pivotTableDefinition16/></ext></extLst>";
            xml += "</pivotTableDefinition>";
            return xml;
        }
        internal Packaging.ZipPackagePart Part
        {
            get;
            set;
        }
        /// <summary>
        /// Individual styles for the pivot table.
        /// </summary>
        public ExcelPivotTableAreaStyleCollection Styles
        {
            get;
            private set;
        }
        /// <summary>
        /// Provides access to the XML data representing the pivottable in the package.
        /// </summary>
        public XmlDocument PivotTableXml { get; private set; }
        /// <summary>
        /// The package internal URI to the pivottable Xml Document.
        /// </summary>
        public Uri PivotTableUri
        {
            get;
            internal set;
        }
        internal Packaging.ZipPackageRelationship Relationship
        {
            get;
            set;
        }
        const string NAME_PATH = "@name";
        const string DISPLAY_NAME_PATH = "@displayName";
        /// <summary>
        /// Name of the pivottable object in Excel
        /// </summary>
        public string Name
        {
            get
            {
                return GetXmlNodeString(NAME_PATH);
            }
            set
            {
                if (WorkSheet.Workbook.ExistsTableName(value))
                {
                    throw (new ArgumentException("PivotTable name is not unique"));
                }
                string prevName = Name;
                if (WorkSheet.Tables._tableNames.ContainsKey(prevName))
                {
                    int ix = WorkSheet.Tables._tableNames[prevName];
                    WorkSheet.Tables._tableNames.Remove(prevName);
                    WorkSheet.Tables._tableNames.Add(value, ix);
                }
                SetXmlNodeString(NAME_PATH, value);
                SetXmlNodeString(DISPLAY_NAME_PATH, CleanDisplayName(value));
            }
        }
        /// <summary>
        /// Reference to the pivot table cache definition object
        /// </summary>
        public ExcelPivotCacheDefinition CacheDefinition
        {
            get;
            private set;
        }
        public bool IsCalculated
        {
            get;
            private set;
        } = false;

        internal List<Dictionary<int[], HashSet<int[]>>> Keys = null;
        internal List<PivotCalculationStore> CalculatedItems = null;
        internal Dictionary<string, PivotCalculationStore> CalculatedFieldReferencedItems = null;
        internal Dictionary<string, PivotCalculationStore> CalculatedFieldRowColumnSubTotals = null;
        internal HashSet<int[]> _rowItems = null;
        internal HashSet<int[]> _colItems = null;
        /// <summary>
        /// Calculates the pivot table 
        /// </summary>
        /// <param name="refreshCache"></param>
        public void Calculate(bool refreshCache = false)
        {
            if (refreshCache || CacheDefinition._cacheReference.Records.RecordCount == 0)
            {
                CacheDefinition.Refresh();
            }
            PivotTableCalculation.Calculate(this, out CalculatedItems, out Keys);
            IsCalculated = true;
        }
        /// <summary>
        /// Returns the calculated grand total for the pivot table. This function works similar to the GetPivotData function used in formulas.
        /// If the pivot table is created in EPPlus without refreshing the cache, the cache will be created.
        /// Please note the any source data containing formulas must be calculated before the pivot table is calculated.
        /// <seealso cref="Calculate(bool)"/>
        /// <seealso cref="IsCalculated"/>
        /// <seealso cref="ExcelPivotCacheDefinition.Refresh"/>
        /// </summary>
        /// <param name="dataFieldName">The name of the data field. If a data field with the name does exist in the table, a #REF! error is returned-</param>
        /// <returns>The calculated value</returns>
        public object GetPivotData(string dataFieldName)
        {
            return GetPivotData(dataFieldName, new List<PivotDataFieldItemSelection>());
        }
        /// <summary>
        /// Returns a calculated value for a row or column field. This function works similar to the GetPivotData function.
        /// If a row or column field is omitted, the subtotal for that field is retrieved.
        /// If the pivot table is not calculated a calculation will be performed without refreshing the pivot cache.
        /// If the pivot table is created in EPPlus without refreshing the cache, the cache will be created.
        /// Please note the any source data containing formulas must be calculated before the pivot table is calculated.
        /// <seealso cref="Calculate(bool)"/>
        /// <seealso cref="IsCalculated"/>
        /// <seealso cref="ExcelPivotCacheDefinition.Refresh"/>
        /// </summary>
        /// <param name="fieldItemSelection">A list of criterias to determin which value to retrieve. If the fieldItemSelection does not exist in the pivot tabvle a #REF! error is returned.</param>
        /// <param name="dataFieldName">The name of the data field. If a data field with the name does exist in the table, a #REF! error is returned-</param>
        /// <returns>The calculated value</returns>
        public object GetPivotData(string dataFieldName, IList<PivotDataFieldItemSelection> fieldItemSelection)
        {
            var dataField = DataFields[dataFieldName];
            if (dataField == null)
            {
                return ErrorValues.RefError;
            }

            if (IsCalculated == false)
            {
                Calculate();
            }

            var items = CacheDefinition._cacheReference.Records.CacheItems;

            var keyFieldIndex = RowColumnFieldIndicies;
            var key = new int[keyFieldIndex.Count];
            var functionFieldIx = -1;
            var function = eSubTotalFunctions.None;
            for (int i = 0; i < keyFieldIndex.Count; i++)
            {
                key[i] = PivotCalculationStore.SumLevelValue;
                for (int j = 0; j < fieldItemSelection.Count; j++)
                {
                    var field = Fields[fieldItemSelection[j].FieldName];

                    if (field == null || (field.IsColumnField==false && field.IsRowField==false))
                    {
                        return ErrorValues.RefError;
                    }
                    if (fieldItemSelection[j].SubtotalFunction != eSubTotalFunctions.None && function != fieldItemSelection[j].SubtotalFunction)
                    {
                        if (function != eSubTotalFunctions.None && (functionFieldIx != field.Index))
                        {
                            return ErrorValues.RefError;
                        }
                        functionFieldIx = field.Index;
                        function = fieldItemSelection[j].SubtotalFunction;
                    }

                    if (field.Index == keyFieldIndex[i])
                    {
                        var cache = field.GetLookup();

                        var isGrouping = field.Grouping != null;

                        if (isGrouping)
                        {
                            var errorValue = GetGroupingKey(fieldItemSelection, ref key, i, j, cache);
                            if (errorValue != null)
                            {
                                return errorValue;
                            }
                        }
                        else
                        {
                            var v = fieldItemSelection[j].Value;
                            if (cache.ContainsKey(v))
                            {
                                key[i] = cache[v];
                            }
                            else
                            {
                                return ErrorValues.RefError;
                            }
                        }
                        break;
                    }
                }
            }

            var dfIx = DataFields.IndexOf(dataField);
            if (PivotTableCalculation.IsReferencingUngroupableKey(key, dataField.Field.PivotTable.RowFields.Count))
            {
                if (Keys[dfIx].TryGetValue(key, out HashSet<int[]> uniqueItems))
                {
                    if (GetMatchingCount(uniqueItems, key, dataField.Field.PivotTable.RowFields.Count, out int[] newKey) == 1)
                    {
                        //key = PivotTableCalculation.GetKeyWithParentLevel(key, newKey, dataField.Field.PivotTable.RowFields.Count);
                        key = newKey;
                    }
                    else
                    {
                        return ErrorValues.RefError;
                    }
                }
                else
                {
                    var sumKey = GetSumKey(key);
                    if (Keys[dfIx].TryGetValue(sumKey, out uniqueItems))
                    {
                        if (GetMatchingCount(uniqueItems, key, RowFields.Count, out int[] newKey) == 1)
                        {
                            //key = PivotTableCalculation.GetKeyWithParentLevel(key, newKey, dataField.Field.PivotTable.RowFields.Count);
                            key = newKey;
                        }
                        else
                        {
                            return ErrorValues.RefError;
                        }
                    }
                    else
                    {
                        return ErrorValues.RefError;
                    }
                }
            }

            //Check if the cell is collapsed or if a subtotal function is set to none and the subtotal is not the lowest collapsed row/column.
            if (IsCollapsedOrNoSubTotalFunction(key, keyFieldIndex, dataField))
            {
                return ErrorValues.RefError;
            }

            if (function == eSubTotalFunctions.None)
            {
                if (CalculatedItems[dfIx].TryGetValue(key, out var value))
                {
                    return value;
                }
            }
            else
            {
                var subTotalKey = $"{functionFieldIx},{dfIx},{function}";
                if (CalculatedFieldRowColumnSubTotals[subTotalKey].TryGetValue(key, out var value))
                {
                    return value;
                }
            }
            if (ExistsValueInTable(key, dfIx))
            {
                return 0D;
            }
            return ErrorValues.RefError;
        }

        private int GetMatchingCount(HashSet<int[]> uniqueItems, int[] key, int colStartIx, out int[] newKey)
        {
            var count = 0;
            newKey = default;
            foreach (var uniqueKey in uniqueItems)
            {
                var equal = true;
                for (int i = 0; i < key.Length; i++)
                {
                    if (key[i] != PivotCalculationStore.SumLevelValue && uniqueKey[i] != key[i])
                    {
                        equal = false;
                        break;
                    }
                }
                if (equal)
                {
                    newKey = (int[])key.Clone();
                    if (PivotKeyUtil.IsRowGrandTotal(key, colStartIx) == false)
                    {
                        var isSum = false;
                        for (int i = 0; i < colStartIx; i++)
                        {
                            if (newKey[i] == PivotCalculationStore.SumLevelValue)
                            {
                                newKey[i] = uniqueKey[i];
                                isSum = true;
                            }
                            else
                            {
                                if (isSum) break;
                            }
                        }
                    }
                    if (PivotKeyUtil.IsColumnGrandTotal(key, colStartIx) == false)
                    {
                        var isSum = false;
                        for (int i = colStartIx; i < key.Length; i++)
                        {
                            if (newKey[i] == PivotCalculationStore.SumLevelValue)
                            {
                                newKey[i] = uniqueKey[i];
                                isSum = true;
                            }
                            else
                            {
                                if (isSum) break;
                            }
                        }
                    }
                    if (count == 1) return 2;
                    count++;
                }
            }
            return count;
        }

        private int[] GetSumKey(int[] key)
        {
            var newKey = (int[])key.Clone();
            var hasSumKey = false;
            for (int i = 0; i < key.Length; i++)
            {
                if (i == RowFields.Count) hasSumKey = false;
                if (newKey[i] == PivotCalculationStore.SumLevelValue)
                {
                    hasSumKey = true;
                }
                if (hasSumKey == true && key[i] != PivotCalculationStore.SumLevelValue)
                {
                    newKey[i] = PivotCalculationStore.SumLevelValue;
                }
            }
            return newKey;
        }

        /// <summary>
        /// Access to the calculated data when the pivot table has been calculated.
        /// <seealso cref="Calculate(bool)"/>
        /// <seealso cref="IsCalculated"/>
        /// <seealso cref="GetPivotData(string, IList{PivotDataFieldItemSelection})"/>
        /// </summary>
        public ExcelPivotTableCalculatedData CalculatedData
        {
            get
            {
                return new ExcelPivotTableCalculatedData(this);
            }
        }

        private bool IsCollapsedOrNoSubTotalFunction(int[] key, List<int> keyFieldIndex, ExcelPivotTableDataField datafield)
        {
            if (IsCollapsedOrNoSubTotalFunction(key, keyFieldIndex, 0, RowFields.Count, datafield))
            {
                return true;
            }
            return IsCollapsedOrNoSubTotalFunction(key, keyFieldIndex, RowFields.Count, key.Length, datafield);
        }

        private bool IsCollapsedOrNoSubTotalFunction(int[] key, List<int> keyFieldIndex, int fromIndex, int toIndex, ExcelPivotTableDataField datafield)
        {
            var isCollapsed = false;
            var isParentFunctionNone = false;
            for (int i = fromIndex; i < toIndex; i++)
            {
                if (key[i] == PivotCalculationStore.SumLevelValue)
                {
                    if (isParentFunctionNone && isCollapsed == false && Keys[DataFields.IndexOf(datafield)].ContainsKey(key) && Keys[DataFields.IndexOf(datafield)][key].Count > 1)
                    {
                        return true;
                    }
                    continue;
                }
                if (isCollapsed) return true;
                var field = Fields[keyFieldIndex[i]];
                var item = field.Items.GetByCacheIndex(key[i]);
                if (item != null && item.ShowDetails == false)
                {
                    isCollapsed = true;
                }
                isParentFunctionNone =
                        field.SubTotalFunctions == eSubTotalFunctions.None ||
                       (field.SubTotalFunctions != eSubTotalFunctions.Default && (field.SubTotalFunctions & GetSubTotalEnum(datafield.Function)) != 0);
            }
            return false;
        }

        private eSubTotalFunctions GetSubTotalEnum(DataFieldFunctions function)
        {
            switch (function)
            {
                case DataFieldFunctions.Sum:
                    return eSubTotalFunctions.Sum;
                case DataFieldFunctions.Average:
                    return eSubTotalFunctions.Avg;
                case DataFieldFunctions.Count:
                    return eSubTotalFunctions.Count;
                case DataFieldFunctions.CountNums:
                    return eSubTotalFunctions.CountA;
                case DataFieldFunctions.Product:
                    return eSubTotalFunctions.Product;
                case DataFieldFunctions.Var:
                    return eSubTotalFunctions.Var;
                case DataFieldFunctions.VarP:
                    return eSubTotalFunctions.VarP;
                case DataFieldFunctions.StdDev:
                    return eSubTotalFunctions.StdDev;
                case DataFieldFunctions.StdDevP:
                    return eSubTotalFunctions.StdDevP;
                case DataFieldFunctions.Min:
                    return eSubTotalFunctions.Min;
                case DataFieldFunctions.Max:
                    return eSubTotalFunctions.Max;
                default:
                    return eSubTotalFunctions.Default;
            }
        }

        private bool ExistsValueInTable(int[] key, int dfIx)
        {
            var rowKey = PivotKeyUtil.GetRowTotalKey(key, RowFields.Count);
            var colKey = PivotKeyUtil.GetColumnTotalKey(key, RowFields.Count);
            return CalculatedItems[dfIx].ContainsKey(rowKey) && CalculatedItems[dfIx].ContainsKey(colKey);
        }

        private ExcelErrorValue GetGroupingKey(IList<PivotDataFieldItemSelection> criteria, ref int[] key, int i, int j, Dictionary<object, int> cache)
        {
            var field = Fields[criteria[j].FieldName];
            if (field.Grouping is ExcelPivotTableFieldNumericGroup grp)
            {
                var s = (criteria[j].Value ?? "").ToString();
                if (s == "<")
                {
                    key[i] = -1;
                }
                else if (s == ">")
                {
                    key[i] = int.MaxValue - 1;
                }
                else
                {
                    var n = ConvertUtil.GetValueDouble(criteria[j].Value);
                    if (n == grp.End)
                    {
                        key[i] = int.MaxValue - 1;
                    }
                    else if (Math.Round(n % grp.Interval, 12) == 0 || Math.Round(n % grp.Interval, 12) == grp.Interval)
                    {
                        key[i] = Convert.ToInt32((n - grp.Start) / grp.Interval);
                    }
                    else
                    {
                        return ErrorValues.RefError;
                    }
                }
            }
            else if (field.DateGrouping == eDateGroupBy.Years)
            {
                var sy = criteria[j].Value.ToString();

                if (cache.ContainsKey(sy))
                {
                    key[i] = cache[sy];
                }
                else
                {
                    return ErrorValues.RefError;
                }
            }
            else if (criteria[j].Value is string s)
            {
                if (cache.ContainsKey(s))
                {
                    key[i] = cache[s];
                }
                else
                {
                    return ErrorValues.RefError;
                }
            }
            else
            {
                if (criteria[j].Value == null)
                {
                    if (cache.ContainsKey(ExcelPivotTable.PivotNullValue))
                    {
                        key[i] = cache[ExcelPivotTable.PivotNullValue] - 1;
                    }
                }
                else
                {
                    var d = ConvertUtil.GetValueDouble(criteria[j].Value, true, true);
                    if (double.IsNaN(d))
                    {
                        return ErrorValues.RefError;
                    }
                    else
                    {
                        var index = Convert.ToInt32(d);
                        key[i] = index;
                    }
                }
            }
            return null;
        }

        private string CleanDisplayName(string name)
        {
            return Regex.Replace(name, @"[^\w\.-_]", "_");
        }
        #region "Public Properties"

        /// <summary>
        /// The worksheet where the pivottable is located
        /// </summary>
        public ExcelWorksheet WorkSheet
        {
            get;
            set;
        }
        /// <summary>
        /// The location of the pivot table
        /// </summary>
        public ExcelAddressBase Address
        {
            get;
            internal set;
        }
        /// <summary>
        /// If multiple datafields are displayed in the row area or the column area
        /// </summary>
        public bool DataOnRows
        {
            get
            {
                return GetXmlNodeBool("@dataOnRows");
            }
            set
            {
                SetXmlNodeBool("@dataOnRows", value);
            }
        }
        /// <summary>
        /// The position of the values in the row- or column- fields list. Position is dependent on <see cref="DataOnRows"/>.
        /// If DataOnRows is true then the position is within the <see cref="ColumnFields"/> collection,
        /// a value of false the position is within the <see cref="RowFields" /> collection.
        /// A negative value or a value out of range of the add the "Σ values" field to the end of the collection.
        /// </summary>
        public int ValuesFieldPosition
        {
            get;
            set;
        } = -1;
        /// <summary>
        /// if true apply legacy table autoformat number format properties.
        /// </summary>
        public bool ApplyNumberFormats
        {
            get
            {
                return GetXmlNodeBool("@applyNumberFormats");
            }
            set
            {
                SetXmlNodeBool("@applyNumberFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat border properties
        /// </summary>
        public bool ApplyBorderFormats
        {
            get
            {
                return GetXmlNodeBool("@applyBorderFormats");
            }
            set
            {
                SetXmlNodeBool("@applyBorderFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat font properties
        /// </summary>
        public bool ApplyFontFormats
        {
            get
            {
                return GetXmlNodeBool("@applyFontFormats");
            }
            set
            {
                SetXmlNodeBool("@applyFontFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat pattern properties
        /// </summary>
        public bool ApplyPatternFormats
        {
            get
            {
                return GetXmlNodeBool("@applyPatternFormats");
            }
            set
            {
                SetXmlNodeBool("@applyPatternFormats", value);
            }
        }
        /// <summary>
        /// If true apply legacy table autoformat width/height properties.
        /// </summary>
        public bool ApplyWidthHeightFormats
        {
            get
            {
                return GetXmlNodeBool("@applyWidthHeightFormats");
            }
            set
            {
                SetXmlNodeBool("@applyWidthHeightFormats", value);
            }
        }
        /// <summary>
        /// Show member property information
        /// </summary>
        public bool ShowMemberPropertyTips
        {
            get
            {
                return GetXmlNodeBool("@showMemberPropertyTips");
            }
            set
            {
                SetXmlNodeBool("@showMemberPropertyTips", value);
            }
        }
        /// <summary>
        /// Show the drill indicators
        /// </summary>
        public bool ShowCalcMember
        {
            get
            {
                return GetXmlNodeBool("@showCalcMbrs");
            }
            set
            {
                SetXmlNodeBool("@showCalcMbrs", value);
            }
        }
        /// <summary>
        /// If the user is prevented from drilling down on a PivotItem or aggregate value
        /// </summary>
        public bool EnableDrill
        {
            get
            {
                return GetXmlNodeBool("@enableDrill", true);
            }
            set
            {
                SetXmlNodeBool("@enableDrill", value);
            }
        }
        /// <summary>
        /// Show the drill down buttons
        /// </summary>
        public bool ShowDrill
        {
            get
            {
                return GetXmlNodeBool("@showDrill", true);
            }
            set
            {
                SetXmlNodeBool("@showDrill", value);
            }
        }
        /// <summary>
        /// If the tooltips should be displayed for PivotTable data cells.
        /// </summary>
        public bool ShowDataTips
        {
            get
            {
                return GetXmlNodeBool("@showDataTips", true);
            }
            set
            {
                SetXmlNodeBool("@showDataTips", value, true);
            }
        }
        /// <summary>
        /// If the row and column titles from the PivotTable should be printed.
        /// </summary>
        public bool FieldPrintTitles
        {
            get
            {
                return GetXmlNodeBool("@fieldPrintTitles");
            }
            set
            {
                SetXmlNodeBool("@fieldPrintTitles", value);
            }
        }
        /// <summary>
        /// If the row and column titles from the PivotTable should be printed.
        /// </summary>
        public bool ItemPrintTitles
        {
            get
            {
                return GetXmlNodeBool("@itemPrintTitles");
            }
            set
            {
                SetXmlNodeBool("@itemPrintTitles", value);
            }
        }
        /// <summary>
        /// If the grand totals should be displayed for the PivotTable columns
        /// </summary>
        public bool ColumnGrandTotals
        {
            get
            {
                return GetXmlNodeBool("@colGrandTotals");
            }
            set
            {
                SetXmlNodeBool("@colGrandTotals", value);
            }
        }
        /// <summary>
        /// If the grand totals should be displayed for the PivotTable rows
        /// </summary>
        public bool RowGrandTotals
        {
            get
            {
                return GetXmlNodeBool("@rowGrandTotals");
            }
            set
            {
                SetXmlNodeBool("@rowGrandTotals", value);
            }
        }
        /// <summary>
        /// If the drill indicators expand collapse buttons should be printed.
        /// </summary>
        public bool PrintDrill
        {
            get
            {
                return GetXmlNodeBool("@printDrill");
            }
            set
            {
                SetXmlNodeBool("@printDrill", value);
            }
        }
        /// <summary>
        /// Indicates whether to show error messages in cells.
        /// </summary>
        public bool ShowError
        {
            get
            {
                return GetXmlNodeBool("@showError");
            }
            set
            {
                SetXmlNodeBool("@showError", value);
            }
        }
        /// <summary>
        /// The string to be displayed in cells that contain errors.
        /// </summary>
        public string ErrorCaption
        {
            get
            {
                return GetXmlNodeString("@errorCaption");
            }
            set
            {
                SetXmlNodeString("@errorCaption", value);
            }
        }
        /// <summary>
        /// Specifies the name of the value area field header in the PivotTable. 
        /// This caption is shown when the PivotTable when two or more fields are in the values area.
        /// </summary>
        public string DataCaption
        {
            get
            {
                return GetXmlNodeString("@dataCaption");
            }
            set
            {
                SetXmlNodeString("@dataCaption", value);
            }
        }
        /// <summary>
        /// Show field headers
        /// </summary>
        public bool ShowHeaders
        {
            get
            {
                return GetXmlNodeBool("@showHeaders");
            }
            set
            {
                SetXmlNodeBool("@showHeaders", value);
            }
        }
        /// <summary>
        /// The number of page fields to display before starting another row or column
        /// </summary>
        public int PageWrap
        {
            get
            {
                return GetXmlNodeInt("@pageWrap");
            }
            set
            {
                if (value < 0)
                {
                    throw new Exception("Value can't be negative");
                }
                SetXmlNodeString("@pageWrap", value.ToString());
            }
        }
        /// <summary>
        /// A boolean that indicates whether legacy auto formatting has been applied to the PivotTable view
        /// </summary>
        public bool UseAutoFormatting
        {
            get
            {
                return GetXmlNodeBool("@useAutoFormatting");
            }
            set
            {
                SetXmlNodeBool("@useAutoFormatting", value);
            }
        }
        /// <summary>
        /// A boolean that indicates if the in-grid drop zones should be displayed at runtime, and if classic layout is applied
        /// </summary>
        public bool GridDropZones
        {
            get
            {
                return GetXmlNodeBool("@gridDropZones");
            }
            set
            {
                SetXmlNodeBool("@gridDropZones", value);
            }
        }
        /// <summary>
        /// The indentation increment for compact axis and can be used to set the Report Layout to Compact Form
        /// </summary>
        public int Indent
        {
            get
            {
                return GetXmlNodeInt("@indent");
            }
            set
            {
                SetXmlNodeString("@indent", value.ToString());
            }
        }
        /// <summary>
        /// A boolean that indicates whether data fields in the PivotTable should be displayed in outline form
        /// </summary>
        public bool OutlineData
        {
            get
            {
                return GetXmlNodeBool("@outlineData");
            }
            set
            {
                SetXmlNodeBool("@outlineData", value);
            }
        }
        /// <summary>
        /// A boolean that indicates whether new fields should have their outline flag set to true
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
        /// A boolean that indicates if the fields of a PivotTable can have multiple filters set on them
        /// </summary>
        public bool MultipleFieldFilters
        {
            get
            {
                return GetXmlNodeBool("@multipleFieldFilters");
            }
            set
            {
                SetXmlNodeBool("@multipleFieldFilters", value);
            }
        }
        /// <summary>
        /// A boolean that indicates if new fields should have their compact flag set to true
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
        /// Sets all pivot table fields <see cref="ExcelPivotTableField.Compact"/> property to the value supplied.
        /// </summary>
        /// <param name="value">The the value for the Compact property.</param>
        public void SetCompact(bool value = true)
        {
            Compact = value;
            foreach (var f in Fields)
            {
                f.Compact = value;
            }
        }
        /// <summary>
        /// A boolean that indicates if the field next to the data field in the PivotTable should be displayed in the same column of the spreadsheet.
        /// </summary>
        public bool CompactData
        {
            get
            {
                return GetXmlNodeBool("@compactData");
            }
            set
            {
                SetXmlNodeBool("@compactData", value);
            }
        }
        /// <summary>
        /// Specifies the string to be displayed for grand totals.
        /// </summary>
        public string GrandTotalCaption
        {
            get
            {
                return GetXmlNodeString("@grandTotalCaption");
            }
            set
            {
                SetXmlNodeString("@grandTotalCaption", value);
            }
        }
        /// <summary>
        /// The text to be displayed in row header in compact mode.
        /// </summary>
        public string RowHeaderCaption
        {
            get
            {
                return GetXmlNodeString("@rowHeaderCaption");
            }
            set
            {
                SetXmlNodeString("@rowHeaderCaption", value);
            }
        }
        /// <summary>
        /// The text to be displayed in column header in compact mode.
        /// </summary>
        public string ColumnHeaderCaption
        {
            get
            {
                return GetXmlNodeString("@colHeaderCaption");
            }
            set
            {
                SetXmlNodeString("@colHeaderCaption", value);
            }
        }
        /// <summary>
        /// The text to be displayed in cells with no value
        /// </summary>
        public string MissingCaption
        {
            get
            {
                return GetXmlNodeString("@missingCaption");
            }
            set
            {
                SetXmlNodeString("@missingCaption", value);
            }
        }
        ExcelPivotTableFilterCollection _filters = null;
        /// <summary>
        /// Filters applied to the pivot table
        /// </summary>
        public ExcelPivotTableFilterCollection Filters
        {
            get
            {
                if (_filters == null)
                {
                    _filters = new ExcelPivotTableFilterCollection(this);
                }
                return _filters;
            }
        }
        const string FIRSTHEADERROW_PATH = "d:location/@firstHeaderRow";
        /// <summary>
        /// The first row of the PivotTable header, relative to the top left cell in the ref value
        /// </summary>
        public int FirstHeaderRow
        {
            get
            {
                return GetXmlNodeInt(FIRSTHEADERROW_PATH);
            }
            set
            {
                SetXmlNodeString(FIRSTHEADERROW_PATH, value.ToString());
            }
        }
        const string FIRSTDATAROW_PATH = "d:location/@firstDataRow";
        /// <summary>
        /// The first column of the PivotTable data, relative to the top left cell in the range
        /// </summary>
        public int FirstDataRow
        {
            get
            {
                return GetXmlNodeInt(FIRSTDATAROW_PATH);
            }
            set
            {
                SetXmlNodeString(FIRSTDATAROW_PATH, value.ToString());
            }
        }
        const string FIRSTDATACOL_PATH = "d:location/@firstDataCol";
        /// <summary>
        /// The first column of the PivotTable data, relative to the top left cell in the range.
        /// </summary>
        public int FirstDataCol
        {
            get
            {
                return GetXmlNodeInt(FIRSTDATACOL_PATH);
            }
            set
            {
                SetXmlNodeString(FIRSTDATACOL_PATH, value.ToString());
            }
        }
        ExcelPivotTableFieldCollection _fields = null;
        /// <summary>
        /// The fields in the table 
        /// </summary>
        public ExcelPivotTableFieldCollection Fields
        {
            get
            {
                if (_fields == null)
                {
                    _fields = new ExcelPivotTableFieldCollection(this);
                }
                return _fields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _rowFields = null;
        /// <summary>
        /// Row label fields 
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection RowFields
        {
            get
            {
                if (_rowFields == null)
                {
                    _rowFields = new ExcelPivotTableRowColumnFieldCollection(this, "rowFields");
                }
                return _rowFields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _columnFields = null;
        /// <summary>
        /// Column label fields 
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection ColumnFields
        {
            get
            {
                if (_columnFields == null)
                {
                    _columnFields = new ExcelPivotTableRowColumnFieldCollection(this, "colFields");
                }
                return _columnFields;
            }
        }
        ExcelPivotTableDataFieldCollection _dataFields = null;
        /// <summary>
        /// Value fields 
        /// </summary>
        public ExcelPivotTableDataFieldCollection DataFields
        {
            get
            {
                if (_dataFields == null)
                {
                    _dataFields = new ExcelPivotTableDataFieldCollection(this);
                }
                return _dataFields;
            }
        }
        ExcelPivotTableRowColumnFieldCollection _pageFields = null;
        /// <summary>
        /// Report filter fields
        /// </summary>
        public ExcelPivotTableRowColumnFieldCollection PageFields
        {
            get
            {
                if (_pageFields == null)
                {
                    _pageFields = new ExcelPivotTableRowColumnFieldCollection(this, "pageFields");
                }
                return _pageFields;
            }
        }
        const string STYLENAME_PATH = "d:pivotTableStyleInfo/@name";
        /// <summary>
        /// Pivot style name. Used for custom styles
        /// </summary>
        public string StyleName
        {
            get
            {
                return GetXmlNodeString(STYLENAME_PATH);
            }
            set
            {
                if (value.StartsWith("PivotStyle"))
                {
                    try
                    {
                        if (Enum.GetNames(typeof(TableStyles)).Any(x => x.Equals(value.Substring(10, value.Length - 10), StringComparison.OrdinalIgnoreCase)))
                        {
                            _tableStyle = (TableStyles)Enum.Parse(typeof(TableStyles), value.Substring(10, value.Length - 10), true);
                        }
                        else
                        {
                            _tableStyle = TableStyles.Custom;
                        }
                    }
                    catch
                    {
                        _tableStyle = TableStyles.Custom;
                    }
                    try
                    {
                        _pivotTableStyle = (PivotTableStyles)Enum.Parse(typeof(PivotTableStyles), value.Substring(10, value.Length - 10), true);
                    }
                    catch
                    {
                        _pivotTableStyle = PivotTableStyles.Custom;
                    }

                }
                else if (value == "None")
                {
                    _tableStyle = TableStyles.None;
                    _pivotTableStyle = PivotTableStyles.None;
                    value = "";
                }
                else
                {
                    _tableStyle = TableStyles.Custom;
                    _pivotTableStyle = PivotTableStyles.Custom;
                }
                SetXmlNodeString(STYLENAME_PATH, value, true);
            }
        }
        const string SHOWCOLHEADERS_PATH = "d:pivotTableStyleInfo/@showColHeaders";
        /// <summary>
        /// Whether to show column headers for the pivot table.
        /// </summary>
        public bool ShowColumnHeaders
        {
            get
            {
                return GetXmlNodeBool(SHOWCOLHEADERS_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWCOLHEADERS_PATH, value);
            }
        }
        const string SHOWCOLSTRIPES_PATH = "d:pivotTableStyleInfo/@showColStripes";
        /// <summary>
        /// Whether to show column stripe formatting for the pivot table.
        /// </summary>
        public bool ShowColumnStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWCOLSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWCOLSTRIPES_PATH, value);
            }
        }
        const string SHOWLASTCOLUMN_PATH = "d:pivotTableStyleInfo/@showLastColumn";
        /// <summary>
        /// Whether to show the last column for the pivot table.
        /// </summary>
        public bool ShowLastColumn
        {
            get
            {
                return GetXmlNodeBool(SHOWLASTCOLUMN_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWLASTCOLUMN_PATH, value);
            }
        }
        const string SHOWROWHEADERS_PATH = "d:pivotTableStyleInfo/@showRowHeaders";
        /// <summary>
        /// Whether to show row headers for the pivot table.
        /// </summary>
        public bool ShowRowHeaders
        {
            get
            {
                return GetXmlNodeBool(SHOWROWHEADERS_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWROWHEADERS_PATH, value);
            }
        }
        const string SHOWROWSTRIPES_PATH = "d:pivotTableStyleInfo/@showRowStripes";
        /// <summary>
        /// Whether to show row stripe formatting for the pivot table.
        /// </summary>
        public bool ShowRowStripes
        {
            get
            {
                return GetXmlNodeBool(SHOWROWSTRIPES_PATH);
            }
            set
            {
                SetXmlNodeBool(SHOWROWSTRIPES_PATH, value);
            }
        }
        TableStyles _tableStyle = Table.TableStyles.Medium6;
        /// <summary>
        /// The table style. If this property is Custom, the style from the StyleName propery is used.
        /// </summary>
        [Obsolete("Use the PivotTableStyle property for more options")]
        public TableStyles TableStyle
        {
            get
            {
                return _tableStyle;
            }
            set
            {
                _tableStyle = value;
                if (value != TableStyles.Custom)
                {
                    StyleName = "PivotStyle" + value.ToString();
                }
            }
        }
        PivotTableStyles _pivotTableStyle = PivotTableStyles.Medium6;
        /// <summary>
        /// The pivot table style. If this property is Custom, the style from the StyleName propery is used.
        /// </summary>
        public PivotTableStyles PivotTableStyle
        {
            get
            {
                return _pivotTableStyle;
            }
            set
            {
                _pivotTableStyle = value;
                if (value != PivotTableStyles.Custom)
                {
                    //SetXmlNodeString(STYLENAME_PATH, "PivotStyle" + value.ToString());
                    StyleName = "PivotStyle" + value.ToString();
                }
            }
        }
        const string _showValuesRowPath = "d:extLst/d:ext[@uri='" + ExtLstUris.PivotTableDefinitionUri + "']/x14:pivotTableDefinition/@hideValuesRow";
        /// <summary>
        /// If the pivot tables value row is visible or not. 
        /// This property only applies when <see cref="GridDropZones"/> is set to false.
        /// </summary>
        public bool ShowValuesRow
        {
            get
            {
                return !GetXmlNodeBool(_showValuesRowPath);
            }
            set
            {
                var node = GetOrCreateExtLstSubNode(ExtLstUris.PivotTableDefinitionUri, "x14");
                var xh = XmlHelperFactory.Create(NameSpaceManager, node);
                xh.SetXmlNodeBool("x14:pivotTableDefinition/@hideValuesRow", !value);
            }
        }

        #endregion
        #region "Internal Properties"
        internal int CacheId
        {
            get
            {
                return GetXmlNodeInt("@cacheId", 0);
            }
            set
            {
                SetXmlNodeInt("@cacheId", value);
            }
        }

        internal List<int> RowColumnFieldIndicies
        {
            get
            {
                return RowFields.Union(ColumnFields).Select(x => x.Cache.Index).ToList();
            }
        }


        internal int ChangeCacheId(int oldCacheId)
        {
            var newCacheId = WorkSheet.Workbook.GetNewPivotCacheId();
            CacheId = newCacheId;
            CacheDefinition._cacheReference.ExtLstCacheId = newCacheId;
            WorkSheet.Workbook.SetXmlNodeInt($"d:pivotCaches/d:pivotCache[@cacheId={oldCacheId}]/@cacheId", newCacheId);

            return newCacheId;
        }

        #endregion
        int _newFilterId = 0;
        internal int GetNewFilterId()
        {
            return _newFilterId++;
        }
        internal void SetNewFilterId(int value)
        {
            if (value >= _newFilterId)
            {
                _newFilterId = value + 1;
            }
        }

        internal void Save()
        {
            if (DataFields.Count > 1)
            {
                XmlElement parentNode;
                int fields;
                if (DataOnRows == true)
                {
                    parentNode = PivotTableXml.SelectSingleNode("//d:rowFields", NameSpaceManager) as XmlElement;
                    if (parentNode == null)
                    {
                        CreateNode("d:rowFields");
                        parentNode = PivotTableXml.SelectSingleNode("//d:rowFields", NameSpaceManager) as XmlElement;
                    }
                    fields = RowFields.Count;
                }
                else
                {
                    parentNode = PivotTableXml.SelectSingleNode("//d:colFields", NameSpaceManager) as XmlElement;
                    if (parentNode == null)
                    {
                        CreateNode("d:colFields");
                        parentNode = PivotTableXml.SelectSingleNode("//d:colFields", NameSpaceManager) as XmlElement;
                    }
                    fields = ColumnFields.Count;
                }

                if (parentNode.SelectSingleNode("d:field[@ x= \"-2\"]", NameSpaceManager) == null)
                {
                    XmlElement fieldNode = PivotTableXml.CreateElement("field", ExcelPackage.schemaMain);
                    fieldNode.SetAttribute("x", "-2");
                    if (ValuesFieldPosition >= 0 && ValuesFieldPosition < fields)
                    {
                        parentNode.InsertBefore(fieldNode, parentNode.ChildNodes[ValuesFieldPosition]);
                    }
                    else
                    {
                        parentNode.AppendChild(fieldNode);
                    }
                }
            }

            SetXmlNodeString("d:location/@ref", Address.Address);

            foreach (var field in Fields)
            {
                field.SaveToXml();
            }

            foreach (var df in DataFields)
            {
                if (string.IsNullOrEmpty(df.Name))
                {

                    string name;
                    if (df.Function == DataFieldFunctions.None)
                    {
                        name = df.Field.Name; //Name must be set or Excel will crash on rename.                                
                    }
                    else
                    {
                        name = df.Function.ToString() + " of " + df.Field.Name; //Name must be set or Excel will crash on rename.
                    }

                    //Make sure name is unique
                    var newName = name;
                    var i = 2;
                    while (DataFields.ExistsDfName(newName, df))
                    {
                        newName = name + (i++).ToString(CultureInfo.InvariantCulture);
                    }
                    df.Name = newName;
                }
            }

            UpdatePivotTableStyles();
            PivotTableXml.Save(Part.GetStream(FileMode.Create));
        }
        private void UpdatePivotTableStyles()
        {
            foreach (ExcelPivotTableAreaStyle a in Styles)
            {
                a.Conditions.UpdateXml();
            }
        }

        internal void Sort()
        {
            foreach (var field in RowFields.Union(ColumnFields))
            {
                if (field.Sort != eSortType.None)
                {
                    field.Items.Sort(field.Sort);
                }
            }
        }

        internal IList<ExcelPivotTableDataField> GetFieldsToCalculate()
        {
            return DataFields.ToList();
        }

        internal void MatchFieldValuesToIndex()
        {
            foreach (var f in Fields)
            {
                //if(f.IsColumnField || f.IsRowField || f.IsPageField || f.Slicer!=null)
                if (f.Items.Count > 0)
                {
                    f.Items.MatchValueToIndex();
                }
            }
        }

        internal void InitCalculation()
        {
            _colItems = _rowItems = null;

            MatchFieldValuesToIndex();
            Filters.ReloadTable();
            foreach (var field in RowFields.Union(ColumnFields))
            {
                field.Items.InitNewCalculation();
                if (field.Sort != eSortType.None)
                {
                    field.Items.Sort(field.Sort);
                }
            }
        }        
        internal List<int[]> GetTableKeys()
        {
            var l = new List<int[]>();
            var rowItems = GetTableRowKeys();
            var colItems = GetTableColumnKeys();
            var keyLength = RowFields.Count + ColumnFields.Count;
            var colStartIx = RowFields.Count;

            if (rowItems.Count == 0)
            {
                for (int c = 0; c < colItems.Count; c++)
                {
                    var currentKey = new int[keyLength];
                    for (var i = 0; i < keyLength; i++)
                    {
                            currentKey[i] = colItems[c][i];
                    }
                    l.Add(currentKey);
                }
            }
            else if (colItems.Count == 0)
            {
                for (int r = 0; r < rowItems.Count; r++)
                {
                    var currentKey = new int[keyLength];
                    for (var i = 0; i < keyLength; i++)
                    {
                        currentKey[i] = rowItems[r][i];
                    }
                    l.Add(currentKey);
                }
            }
            else
            {
                for (int r = 0; r < rowItems.Count; r++)
                {
                    for (int c = 0; c < colItems.Count; c++)
                    {
                        var currentKey = new int[keyLength];
                        for (var i = 0; i < keyLength; i++)
                        {
                            if (i < colStartIx)
                            {
                                currentKey[i] = rowItems[r][i];
                            }
                            else
                            {
                                currentKey[i] = colItems[c][i - colStartIx];
                            }
                        }
                        l.Add(currentKey);
                    }
                }
            }
            return l;
        }

        internal List<int[]> GetTableRowKeys()
        {
            return _rowItems.OrderBy(x=>x, ArrayComparer.Instance).ToList<int[]>();
        }
        internal List<int[]> GetTableColumnKeys()
        {
            return _colItems.OrderBy(x => x, ArrayComparer.Instance).ToList<int[]>();
        }
    }
}
