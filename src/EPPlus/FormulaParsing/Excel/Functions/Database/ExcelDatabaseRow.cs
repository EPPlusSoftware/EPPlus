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
using System.Linq;
using System.Text;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace OfficeOpenXml.FormulaParsing.Excel.Functions.Database
{
    /// <summary>
    /// Database row
    /// </summary>
    internal class ExcelDatabaseRow
    {
        private Dictionary<int, string> _fieldIndexes = new Dictionary<int, string>();
        private readonly Dictionary<string, object> _items = new Dictionary<string, object>();
        private int _colIndex = 1;
        /// <summary>
        /// Get object at field
        /// </summary>
        /// <param name="field"></param>
        /// <returns></returns>
        public object this[string field]
        {
            get { return _items[field]; }

            set
            {
                _items[field] = value;
                _fieldIndexes[_colIndex++] = field;
            }
        }
        /// <summary>
        /// Fetch field from indexes then return that field from within the row
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public object this[int index]
        {
            get
            {
                var field = _fieldIndexes[index];
                return _items[field];
            }
        }
        
        
    }
}
