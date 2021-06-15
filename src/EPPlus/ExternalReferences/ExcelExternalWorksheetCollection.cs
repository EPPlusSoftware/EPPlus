/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  04/16/2021         EPPlus Software AB       EPPlus 5.7
 *************************************************************************************************/
using OfficeOpenXml.Core;
using System.Collections.Generic;

namespace OfficeOpenXml.ExternalReferences
{
    /// <summary>
    /// A collection of external worksheets
    /// </summary>
    public class ExcelExternalWorksheetCollection : EPPlusReadOnlyList<ExcelExternalWorksheet>
    {
        Dictionary<string, int> _sheetNames=new Dictionary<string, int>();
        /// <summary>
        /// The indexer to reference the external worksheet objects 
        /// </summary>
        /// <param name="name">The name of the worksheet</param>
        /// <returns></returns>
        public ExcelExternalWorksheet this[string name]
        {
            get
            {
                if (_sheetNames.ContainsKey(name))
                {
                    return _list[_sheetNames[name]];
                }
                return null;
            }
        }
        internal override void Add(ExcelExternalWorksheet item)
        {
            _sheetNames.Add(item.Name, _list.Count);
            base.Add(item);
        }
    }
}
