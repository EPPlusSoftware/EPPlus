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
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromDictionaries : LoadFunctionBase
    {
#if !NET35 && !NET40
        public LoadFromDictionaries(ExcelRangeBase range, IEnumerable<dynamic> items, LoadFromDictionariesParams parameters)
            : this(range, ConvertToDictionaries(items), parameters)
        {

        }
#endif

        public LoadFromDictionaries(ExcelRangeBase range, IEnumerable<IDictionary<string, object>> items, LoadFromDictionariesParams parameters) 
            : base(range, parameters)
        {
            _items = items;
            _keys = parameters.Keys;
            _headerParsingType = parameters.HeaderParsingType;
            _cultureInfo = parameters.Culture ?? CultureInfo.CurrentCulture;
            if (items == null || !items.Any())
            {
                _keys = Enumerable.Empty<string>();
            }
            else
            {
                var firstItem = items.First();
                if (_keys == null || !_keys.Any())
                {
                    _keys = firstItem.Keys;
                }
                else
                {
                    _keys = parameters.Keys;
                }
            }
            _dataTypes = parameters.DataTypes ?? new eDataTypes[0];
        }

        private readonly IEnumerable<IDictionary<string, object>> _items;
        private readonly IEnumerable<string> _keys;
        private readonly eDataTypes[] _dataTypes;
        private readonly HeaderParsingTypes _headerParsingType;
        private readonly CultureInfo _cultureInfo;

#if !NET35 && !NET40
        private static IEnumerable<IDictionary<string, object>> ConvertToDictionaries(IEnumerable<dynamic> items)
        {
            var result = new List<Dictionary<string, object>>();
            foreach(var item in items)
            {
                var obj = item as object;
                if(obj != null)
                {
                    var dict = new Dictionary<string, object>();
                    var members = obj.GetType().GetMembers();
                    foreach(var member in members)
                    {
                        var key = member.Name;
                        object value = null;
                        if (member is PropertyInfo)
                        {
                            value = ((PropertyInfo)member).GetValue(obj);
                            dict.Add(key, value);
                        }
                        else if (member is FieldInfo)
                        {
                            value = ((FieldInfo)member).GetValue(obj);
                            dict.Add(key, value);
                        }
                        
                    }
                    if(dict.Count > 0)
                    {
                        result.Add(dict);
                    }
                }
            }
            return result;
        }

#endif        

        protected override void LoadInternal(object[,] values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int, string> columnFormats)
        {
            columnFormats = new Dictionary<int, string>();
            formulaCells = new Dictionary<int, FormulaCell>();
            int col = 0, row = 0;
            if (PrintHeaders && _keys.Count() > 0)
            {
                foreach (var key in _keys)
                {
                    if (transpose)
                    {
                        values[row++, col] = ParseHeader(key);
                    }
                    else
                    {
                        values[row, col++] = ParseHeader(key);
                    }
                }
                if (transpose) 
                { 
                    col++;
                }
                else 
                {
                    row++;
                }
            }
            foreach (var item in _items)
            {
                if(transpose)
                {
                    row = 0;
                }
                else
                {
                    col = 0;
                }
                foreach (var key in _keys)
                {
                    if (item.ContainsKey(key))
                    {
                        var v = item[key];
                        var dtCheck = transpose ? row < _dataTypes.Length : col < _dataTypes.Length;
                        if (dtCheck && v != null)
                        {
                            var dataType = _dataTypes[col];
                            switch(dataType)
                            {
                                case eDataTypes.Percent:
                                case eDataTypes.Number:
                                    if(double.TryParse(v.ToString(), NumberStyles.Float | NumberStyles.Number, _cultureInfo, out double d))
                                    {
                                        if(dataType == eDataTypes.Percent)
                                        {
                                            d /= 100d;
                                        }
                                        values[row, col] = d;
                                    }
                                    break;
                                case eDataTypes.DateTime:
                                    if(DateTime.TryParse(v.ToString(), out DateTime dt))
                                    {
                                        values[row, col] = dt;
                                    }
                                    break;
                                case eDataTypes.String:
                                    values[row, col] = v.ToString();
                                    break;
                                default:
                                    values[row, col] = v;
                                    break;

                            }
                        }
                        else
                        {
                            values[row, col] = item[key];
                        }
                        if(transpose)
                        {
                            row++;
                        }
                        else
                        {
                            col++;
                        }
                    }
                    else
                    {
                        if (transpose)
                        {
                            row++;
                        }
                        else
                        {
                            col++;
                        }
                    }
                }
                if (transpose)
                {
                    col++;
                }
                else
                {
                    row++;
                }
            }
        }

        protected override int GetNumberOfRows()
        {
            if (_items == null) return 0;
            return _items.Count();
        }

        protected override int GetNumberOfColumns()
        {
            if (_keys == null) return 0;
            return _keys.Count();
        }

        private string ParseHeader(string header)
        {
            switch (_headerParsingType)
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
