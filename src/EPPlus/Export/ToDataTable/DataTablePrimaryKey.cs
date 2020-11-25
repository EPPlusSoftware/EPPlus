/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/15/2020         EPPlus Software AB       ToDataTable function
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeOpenXml.Export.ToDataTable
{
    internal class DataTablePrimaryKey
    {
        private readonly ToDataTableOptions _options;
        private readonly HashSet<string> _keyNames = new HashSet<string>();

        public DataTablePrimaryKey(ToDataTableOptions options)
        {
            _options = options;
            Initialize();
        }

        private void Initialize()
        {
            if(_options.PrimaryKeyNames.Any())
            {
                foreach(var name in _options.PrimaryKeyNames)
                {
                    AddPrimaryKeyName(name);
                }
            }
            else if(_options.PrimaryKeyIndexes.Any())
            {
                foreach(var ix in _options.PrimaryKeyIndexes)
                {
                    try
                    {
                        var mapping = _options.Mappings.GetByRangeIndex(ix);
                        AddPrimaryKeyName(mapping.DataColumnName);
                    }
                    catch(ArgumentOutOfRangeException e)
                    {
                        throw new ArgumentOutOfRangeException("primary key index out of range: " + ix, e);
                    }
                }
            }
        }

        private void AddPrimaryKeyName(string name)
        {
            if (_keyNames.Contains(name))
            {
                throw new InvalidOperationException("Duplicate primary key name: " + name);
            }
            if (!_options.Mappings.Exists(x => x.DataColumnName == name))
            {
                throw new InvalidOperationException("Invalid primary key name, no corresponding DataColumn: " + name);
            }
            _keyNames.Add(name);
        }

        internal IEnumerable<string> KeyNames => _keyNames;

        internal bool HasKeys => _keyNames.Any();

        internal bool ContainsKey(string key)
        {
            return _keyNames.Contains(key);
        }
    }
}
