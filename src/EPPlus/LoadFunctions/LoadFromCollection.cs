/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/16/2020         EPPlus Software AB       EPPlus 5.2.1
 *************************************************************************************************/
using OfficeOpenXml.Compatibility;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using OfficeOpenXml.LoadFunctions.Params;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromCollection<T> : LoadFunctionBase
    {
        public LoadFromCollection(ExcelRangeBase range, IEnumerable<T> items, LoadFromCollectionParams parameters) : base(range, parameters)
        {
            _items = items;
            _bindingFlags = parameters.BindingFlags;
            _headerParsingType = parameters.HeaderParsingType;
            var type = typeof(T);
            var tableAttr = type.GetFirstAttributeOfType<EpplusTableAttribute>();
            if(tableAttr != null)
            {
                ShowFirstColumn = tableAttr.ShowFirstColumn;
                ShowLastColumn = tableAttr.ShowLastColumn;
                ShowTotal = tableAttr.ShowTotal;
            }
            if (parameters.Members == null)
            {
                var columns = SetupColumns();
                _columns = columns.ToArray();
            }
            else
            {
                _columns = parameters.Members.Select(x => new ColumnInfo { MemberInfo = x }).ToArray();
                if (_columns.Length == 0)   //Fixes issue 15555
                {
                    throw (new ArgumentException("Parameter Members must have at least one property. Length is zero"));
                }
                foreach (var columnInfo in _columns)
                {
                    if (columnInfo.MemberInfo == null) continue;
                    var member = columnInfo.MemberInfo;
                    if (member.DeclaringType != null && member.DeclaringType != type)
                    {
                        _isSameType = false;
                    }

                    //Fixing inverted check for IsSubclassOf / Pullrequest from tomdam
                    if (member.DeclaringType != null && member.DeclaringType != type && !TypeCompat.IsSubclassOf(type, member.DeclaringType) && !TypeCompat.IsSubclassOf(member.DeclaringType, type))
                    {
                        throw new InvalidCastException("Supplied properties in parameter Properties must be of the same type as T (or an assignable type from T)");
                    }
                }
            }
        }

        private readonly BindingFlags _bindingFlags;
        private readonly ColumnInfo[] _columns;
        private readonly HeaderParsingTypes _headerParsingType;
        private readonly IEnumerable<T> _items;
        private readonly bool _isSameType = true;

        protected override int GetNumberOfColumns()
        {
            return _columns.Length == 0 ? 1 : _columns.Length;
        }

        protected override int GetNumberOfRows()
        {
            if (_items == null) return 0;
            return _items.Count();
        }

        protected override void PostProcessTable(ExcelTable table, ExcelRangeBase range)
        {
            for(var ix = 0; ix < table.Columns.Count; ix++)
            {
                if (ix >= _columns.Length) break;
                var totalsRowFormula = _columns[ix].TotalsRowFormula;
                var totalsRowLabel = _columns[ix].TotalsRowLabel;
                if (!string.IsNullOrEmpty(totalsRowFormula))
                {
                    table.Columns[ix].TotalsRowFormula = totalsRowFormula;
                }
                else if(!string.IsNullOrEmpty(totalsRowLabel))
                {
                    table.Columns[ix].TotalsRowLabel = _columns[ix].TotalsRowLabel;
                    table.Columns[ix].TotalsRowFunction = RowFunctions.None;
                }
                else
                {
                    table.Columns[ix].TotalsRowFunction = _columns[ix].TotalsRowFunction;
                }
                
                if(!string.IsNullOrEmpty(_columns[ix].TotalsRowNumberFormat))
                {
                    var row = range._toRow + 1;
                    var col = range._fromCol + _columns[ix].Index;
                    range.Worksheet.Cells[row, col].Style.Numberformat.Format = _columns[ix].TotalsRowNumberFormat;
                }
            }
        }



        protected override void LoadInternal(object[,] values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int, string> columnFormats)
        {

            int col = 0, row = 0;
            columnFormats = new Dictionary<int, string>();
            formulaCells = new Dictionary<int, FormulaCell>();
            if (_columns.Length > 0 && PrintHeaders)
            {
                SetHeaders(values, columnFormats, ref col, ref row);
            }

            if (!_items.Any() && (_columns.Length == 0 || PrintHeaders == false))
            {
                return;
            }

            SetValuesAndFormulas(values, formulaCells, ref col, ref row);
        }

        private List<ColumnInfo> SetupColumns()
        {
            var type = typeof(T);
            var members = type.GetProperties(_bindingFlags);
            var result = new List<ColumnInfo>();
            if (type.HasMemberWithPropertyOfType<EpplusTableColumnAttribute>())
            {
                foreach (var member in members)
                {
                    if (member.HasPropertyOfType<EpplusIgnore>())
                    {
                        continue;
                    }
                    var sortOrder = -1;
                    var numberFormat = string.Empty;
                    var rowFunction = RowFunctions.None;
                    var totalsRowNumberFormat = string.Empty;
                    var totalsRowLabel = string.Empty;
                    var totalsRowFormula = string.Empty;
                    var epplusColumnAttr = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
                    if (epplusColumnAttr != null)
                    {
                        sortOrder = epplusColumnAttr.Order;
                        numberFormat = epplusColumnAttr.NumberFormat;
                        rowFunction = epplusColumnAttr.TotalsRowFunction;
                        totalsRowNumberFormat = epplusColumnAttr.TotalsRowNumberFormat;
                        totalsRowLabel = epplusColumnAttr.TotalsRowLabel;
                        totalsRowFormula = epplusColumnAttr.TotalsRowFormula;
                    }
                    result.Add(new ColumnInfo
                    {
                        SortOrder = sortOrder,
                        MemberInfo = member,
                        NumberFormat = numberFormat,
                        TotalsRowFunction = rowFunction,
                        TotalsRowNumberFormat = totalsRowNumberFormat,
                        TotalsRowLabel = totalsRowLabel,
                        TotalsRowFormula = totalsRowFormula
                    });
                }
                ReindexAndSortColumns(result);
            }
            else
            {
                var index = 0;
                result = members.Select(x => new ColumnInfo { Index = index++, MemberInfo = x }).ToList();
            }
            var formulaColumnAttributes = type.FindAttributesOfType<EpplusFormulaTableColumnAttribute>();
            if (formulaColumnAttributes != null && formulaColumnAttributes.Any())
            {
                foreach (var attr in formulaColumnAttributes)
                {
                    result.Add(new ColumnInfo 
                    { 
                        SortOrder = attr.Order, 
                        Header = attr.Header, 
                        Formula = attr.Formula, 
                        FormulaR1C1 = attr.FormulaR1C1, 
                        NumberFormat = attr.NumberFormat,
                        TotalsRowFunction = attr.TotalsRowFunction,
                        TotalsRowNumberFormat = attr.TotalsRowNumberFormat
                    });
                }
                ReindexAndSortColumns(result);
            }
            return result;
        }

        private static void ReindexAndSortColumns(List<ColumnInfo> result)
        {
            var index = 0;
            result.Sort((a, b) => a.SortOrder.CompareTo(b.SortOrder));
            result.ForEach(x => x.Index = index++);
        }

        private void SetValuesAndFormulas(object[,] values, Dictionary<int, FormulaCell> formulaCells, ref int col, ref int row)
        {
            var nMembers = GetNumberOfColumns();
            foreach (var item in _items)
            {
                if (item == null)
                {
                    col = GetNumberOfColumns();
                }
                else
                {
                    col = 0;
                    if (item is string || item is decimal || item is DateTime || TypeCompat.IsPrimitive(item))
                    {
                        values[row, col++] = item;
                    }
                    else
                    {
                        foreach (var colInfo in _columns)
                        {
                            if (colInfo.MemberInfo != null)
                            {
                                var member = colInfo.MemberInfo;
                                if (_isSameType == false && item.GetType().GetMember(member.Name, _bindingFlags).Length == 0)
                                {
                                    col++;
                                    continue; //Check if the property exists if and inherited class is used
                                }
                                else if (member is PropertyInfo)
                                {
                                    values[row, col++] = ((PropertyInfo)member).GetValue(item, null);
                                }
                                else if (member is FieldInfo)
                                {
                                    values[row, col++] = ((FieldInfo)member).GetValue(item);
                                }
                                else if (member is MethodInfo)
                                {
                                    values[row, col++] = ((MethodInfo)member).Invoke(item, null);
                                }
                            }
                            else if (!string.IsNullOrEmpty(colInfo.Formula))
                            {
                                formulaCells[colInfo.Index] = new FormulaCell { Formula = colInfo.Formula };
                            }
                            else if (!string.IsNullOrEmpty(colInfo.FormulaR1C1))
                            {
                                formulaCells[colInfo.Index] = new FormulaCell { FormulaR1C1 = colInfo.FormulaR1C1 };
                            }
                        }
                    }
                }
                row++;
            }
        }

        private void SetHeaders(object[,] values, Dictionary<int, string> columnFormats, ref int col, ref int row)
        {
            foreach (var colInfo in _columns)
            {
                var header = string.Empty;
                if (colInfo.MemberInfo != null)
                {
                    // column data based on a property read with reflection
                    var member = colInfo.MemberInfo;
                    var epplusColumnAttribute = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
                    if (epplusColumnAttribute != null)
                    {
                        if (!string.IsNullOrEmpty(epplusColumnAttribute.Header))
                        {
                            header = epplusColumnAttribute.Header;
                        }
                        else
                        {
                            header = ParseHeader(member.Name);
                        }
                        if (!string.IsNullOrEmpty(epplusColumnAttribute.NumberFormat))
                        {
                            columnFormats.Add(col, epplusColumnAttribute.NumberFormat);
                        }
                    }
                    else
                    {
                        var descriptionAttribute = member.GetFirstAttributeOfType<DescriptionAttribute>();
                        if (descriptionAttribute != null)
                        {
                            header = descriptionAttribute.Description;
                        }
                        else
                        {
                            var displayNameAttribute = member.GetFirstAttributeOfType<DisplayNameAttribute>();
                            if (displayNameAttribute != null)
                            {
                                header = displayNameAttribute.DisplayName;
                            }
                            else
                            {
                                header = ParseHeader(member.Name);
                            }
                        }
                    }
                }
                else
                {
                    // column is a FormulaColumn
                    header = colInfo.Header;
                    columnFormats.Add(colInfo.Index, colInfo.NumberFormat);
                }

                values[row, col++] = header;
            }
            row++;
        }

        private string ParseHeader(string header)
        {
            switch(_headerParsingType)
            {
                case HeaderParsingTypes.Preserve:
                    return header;
                case HeaderParsingTypes.UnderscoreToSpace:
                    return header.Replace("_", " ");
                case HeaderParsingTypes.CamelCaseToSpace:
                    return Regex.Replace(header, "([A-Z])", " $1", RegexOptions.Compiled).Trim();
                case HeaderParsingTypes.UnderscoreAndCamelCaseToSpace:
                    header = Regex.Replace(header, "([A-Z])", " $1", RegexOptions.Compiled).Trim();
                    return header.Replace("_ ", "_").Replace("_", " ");
                default:
                    return header;
            }
        }
    }
}

