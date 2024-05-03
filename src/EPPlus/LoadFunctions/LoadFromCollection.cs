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
using System.Text.RegularExpressions;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Utils;
using System.IO;
#if !NET35
using System.ComponentModel.DataAnnotations;
#endif

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromCollection<T> : LoadFunctionBase
    {
        public LoadFromCollection(ExcelRangeBase range, IEnumerable<T> items, LoadFromCollectionParams parameters) : base(range, parameters)
        {
            _items = items;
            _bindingFlags = parameters.BindingFlags;
            _headerParsingType = parameters.HeaderParsingType;
            _numberFormatProvider = parameters.NumberFormatProvider;
            var type = typeof(T);
            var tableAttr = type.GetFirstAttributeOfType<EpplusTableAttribute>();
            if(tableAttr != null)
            {
                ShowFirstColumn = tableAttr.ShowFirstColumn;
                ShowLastColumn = tableAttr.ShowLastColumn;
                ShowTotal = tableAttr.ShowTotal;
            }
            LoadFromCollectionColumns<T> cols;
            if (parameters.Members == null)
            {
                cols = new LoadFromCollectionColumns<T>(parameters);
                var columns = cols.Setup();
                _columns = columns.ToArray();
                SetHiddenColumns();
            }
            else
            {
                if (parameters.Members.Length == 0)   //Fixes issue 15555
                {
                    throw (new ArgumentException("Parameter Members must have at least one property. Length is zero"));
                }
                cols = new LoadFromCollectionColumns<T>(parameters);
                var columns = cols.Setup();
                _columns = columns.ToArray();
            }
        }

        private readonly BindingFlags _bindingFlags;
        private readonly ColumnInfo[] _columns;
        private readonly HeaderParsingTypes _headerParsingType;
        private readonly IEnumerable<T> _items;
        private IExcelNumberFormatProvider _numberFormatProvider;

        internal List<string> SortOrderProperties
        {
            get;
            private set;
        }

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
            if (_columns.Length > 0)
            {
                if(PrintHeaders)
                {
                    SetHeaders(values, columnFormats, ref col, ref row);
                }
                else
                {
                    SetNumberFormats(columnFormats, col);
                }
            }

            if (!_items.Any() && (_columns.Length == 0 || PrintHeaders == false))
            {
                return;
            }
            SetValuesAndFormulas(values, formulaCells, ref col, ref row);
        }

        private void SetHiddenColumns()
        {
            for (var colIx = 0; colIx < _columns.Length; colIx++)
            {
                var columnInfo = _columns[colIx];
                if (columnInfo.Hidden)
                {
                    Range.Worksheet.Column(Range._fromCol + colIx).Hidden = true;
                }
            }
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
                    var t = item.GetType();
                    if (item is string || item is decimal || item is DateTime || t.IsPrimitive)
                    {
                        if (transpose)
                        {
                            values[col++, row] = item;
                        }
                        else
                        {
                            values[row, col++] = item;
                        }
                    }
                    else if(t.IsEnum)
                    {    
                        if(transpose)
                        {
                            values[col++, row] = GetEnumValue(item, t); ;
                        }
                        else
                        {
                            values[row, col++] = GetEnumValue(item, t); ;
                        }
                    }
                    else
                    {
                        foreach (var colInfo in _columns)
                        {
                            object v = null;
                            if (colInfo.Path != null && colInfo.Path.IsFormulaColumn == false && colInfo.Path.Depth > 0)
                            {
                                v = colInfo.Path.GetLastMemberValue(item, _bindingFlags);
#if (!NET35)
                                if (v != null)
                                {
                                    var type = v.GetType();
                                    if (type.IsEnum)
                                    {
                                        v = GetEnumValue(v, type);
                                    }
                                }
#endif
                                if (transpose)
                                {
                                    values[col++, row] = v;
                                }
                                else
                                {
                                    values[row, col++] = v;
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

        private static string GetEnumValue(object item, Type t)
        {
#if (NET35)
            return item.ToString();
#else
            var v = item.ToString();
            var m = t.GetMember(v).FirstOrDefault();
            var da = m.GetCustomAttribute<DescriptionAttribute>();
            return da?.Description ?? v;
#endif            
        }

        private string GetColumnFormatById(int numberFormatId)
        {
            if(_numberFormatProvider == null)
            {
                var attr = typeof(T).GetFirstAttributeOfType<EpplusTableAttribute>();
                if (attr != null && attr.NumberFormatProviderType != null)
                {
                    _numberFormatProvider = Activator.CreateInstance(attr.NumberFormatProviderType) as IExcelNumberFormatProvider;
                    return _numberFormatProvider.GetFormat(numberFormatId);
                }
                else
                {
                    return string.Empty;
                }
            }
            else
            {
                return _numberFormatProvider.GetFormat(numberFormatId);
            }
        }

        private void SetNumberFormats(Dictionary<int, string> columnFormats, int col)
        {
            var column = col;
            foreach (var colInfo in _columns)
            {
                if(colInfo.MemberInfo == null || colInfo.MemberInfo.HasAttributeOfType<EpplusTableColumnAttribute>() == false)
                {
                    continue;
                }
                SetNumberFormatOnColumn(columnFormats, col++, colInfo);

            }

        }

        private void SetNumberFormatOnColumn(Dictionary<int, string> columnFormats, int col, ColumnInfo colInfo)
        {
            if (colInfo.MemberInfo != null && colInfo.IsDictionaryProperty == false)
            {
                var member = colInfo.MemberInfo;
                var epplusColumnAttribute = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();

                if (!string.IsNullOrEmpty(epplusColumnAttribute.NumberFormat))
                {
                    columnFormats.Add(col, epplusColumnAttribute.NumberFormat);
                }
                else if (epplusColumnAttribute.NumberFormatId > int.MinValue)
                {
                    var format = GetColumnFormatById(epplusColumnAttribute.NumberFormatId);
                    if (_numberFormatProvider == null) throw new ArgumentNullException("NumberFormatProvider", "NumberFormatId was set on a column attribute, but no instance of IExcelNumberFormatProvider was supplied to the function. This can be done either via ExcelTableAttribute.NumberFormatProviderType or via the LoadFromCollectionParams.SetNumberFormatProvider method.");
                    if (!string.IsNullOrEmpty(format))
                    {
                        columnFormats.Add(col, format);
                    }
                }
            }
        }

        private void SetHeaders(object[,] values, Dictionary<int, string> columnFormats, ref int col, ref int row)
        {
            foreach (var colInfo in _columns)
            {
                var header = colInfo.Header;
                
                // if the header is already set and contains a space it doesn't need more formatting or validation.
                var useExistingHeader = !string.IsNullOrEmpty(header) && header.Contains(" ");

                if (colInfo.MemberInfo != null && colInfo.IsDictionaryProperty == false)
                {
                    // column data based on a property read with reflection
                    var member = colInfo.MemberInfo;
                    var epplusColumnAttribute = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
                    if (epplusColumnAttribute != null)
                    {
                        if (!useExistingHeader)
                        {
                            if (!string.IsNullOrEmpty(epplusColumnAttribute.Header))
                            {
                                header = epplusColumnAttribute.Header;
                            }
                            else if(string.IsNullOrEmpty(colInfo.Header))
                            {
                                header = ParseHeader(member.Name);
                            }
                        }
                        SetNumberFormatOnColumn(columnFormats, col, colInfo);
                    }
                    else if (!useExistingHeader)
                    {
                        var dotNetHeader = GetHeaderFromDotNetAttributes(member);
                        if (!string.IsNullOrEmpty(dotNetHeader))
                        {
                            header = dotNetHeader;
                        }
                        else if(!string.IsNullOrEmpty(colInfo.Header) && colInfo.Header != member.Name)
                        {
                            header = colInfo.Header;
                        }
                        else
                        {
                            header = ParseHeader(member.Name);
                        }
                    }
                }
                else if(colInfo.IsDictionaryProperty == false)
                {
                    // column is a FormulaColumn
                    header = colInfo.Header;
                    columnFormats.Add(colInfo.Index, colInfo.NumberFormat);
                }
                if(transpose)
                {
                    values[col++, row] = header;
                }
                else
                {
                    values[row, col++] = header;
                }
            }
            row++;
        }

        private string GetHeaderFromDotNetAttributes(MemberInfo member)
        {
            var descriptionAttribute = member.GetFirstAttributeOfType<DescriptionAttribute>();
            if (descriptionAttribute != null)
            {
                return descriptionAttribute.Description;
            }
            var displayNameAttribute = member.GetFirstAttributeOfType<DisplayNameAttribute>();
            if (displayNameAttribute != null)
            {
                return displayNameAttribute.DisplayName;
            }
#if !NET35
            var displayAttribute = member.GetFirstAttributeOfType<DisplayAttribute>();
            if (displayAttribute != null)
            {
                return displayAttribute.Name;
            }
#endif
            return default;
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

