/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  08/286/2021         EPPlus Software AB       EPPlus 5.7.5
 *************************************************************************************************/
using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.LoadFunctions.ReflectionHelpers;
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromCollectionColumns<T>
    {
        public LoadFromCollectionColumns(LoadFromCollectionParams parameters)
            : this(new MemberPathScanner(typeof(T).GetTypeOrUnderlyingType(), parameters))
        { }

        public LoadFromCollectionColumns(MemberPathScanner scanner)
        {
            _scanner = scanner;
        }

        private readonly MemberPathScanner _scanner;

        internal ColumnInfoCollection Setup()
        {
            var result = new ColumnInfoCollection();
            var paths = _scanner.GetPaths();
            foreach (var path in paths)
            { 
                result.Add(new ColumnInfo(path));
            }
            var formulaColumnAttributes = typeof(T).FindAttributesOfType<EpplusFormulaTableColumnAttribute>();
            if (formulaColumnAttributes != null && formulaColumnAttributes.Any())
            {
                foreach (var attr in formulaColumnAttributes)
                {
                    result.Add(new ColumnInfo(attr));
                }
            }
            result.ReindexAndSortColumns();
            return result;
        }
    }
}
