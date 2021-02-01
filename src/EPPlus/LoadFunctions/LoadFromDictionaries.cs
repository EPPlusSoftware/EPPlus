﻿/*************************************************************************************************
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
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromDictionaries : LoadFunctionBase
    {
        public LoadFromDictionaries(ExcelRangeBase range, IEnumerable<IDictionary<string, object>> items, LoadFromDictionariesParams parameters) 
            : base(range, parameters)
        {
            _items = items;
            _keys = parameters.Keys;
            _headerParsingType = parameters.HeaderParsingType;
            if(items == null || !items.Any())
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
        }

        private readonly IEnumerable<IDictionary<string, object>> _items;
        private readonly IEnumerable<string> _keys;
        private readonly HeaderParsingTypes _headerParsingType;

        

        protected override void LoadInternal(object[,] values, out Dictionary<int, FormulaCell> formulaCells, out Dictionary<int, string> columnFormats)
        {
            columnFormats = new Dictionary<int, string>();
            formulaCells = new Dictionary<int, FormulaCell>();
            int col = 0, row = 0;
            if (PrintHeaders && _keys.Count() > 0)
            {
                foreach (var key in _keys)
                {
                    values[row, col++] = ParseHeader(key);
                }
                row++;
            }
            foreach (var item in _items)
            {
                col = 0;
                foreach (var key in _keys)
                {
                    if (item.ContainsKey(key))
                    {
                        values[row, col++] = item[key];
                    }
                    else
                    {
                        col++;
                    }
                }
                row++;
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
