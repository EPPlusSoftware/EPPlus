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
using OfficeOpenXml.Table;
using OfficeOpenXml.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    internal class LoadFromCollectionColumns<T>
    {
        public LoadFromCollectionColumns(LoadFromCollectionParams parameters)
        {
            _params = parameters;
            _bindingFlags = parameters.BindingFlags;
        }


        private readonly LoadFromCollectionParams _params;
        private readonly BindingFlags _bindingFlags;

        internal ColumnInfoCollection Setup()
        {
            var result = new ColumnInfoCollection();
            var t = typeof(T);
            var ut = Nullable.GetUnderlyingType(t);
            if (ut != null)
            {
                t = ut;
            }

            var memberPathScanner = new MemberPathScanner(t, _params);
            var paths = memberPathScanner.GetPaths();
            foreach (var path in paths)
            {
                var pathItem = path.Last();
                var col = new ColumnInfo
                {
                    Header = path.GetHeader(),
                    Path = path,
                    IsDictionaryProperty = pathItem.IsDictionaryColumn,
                    MemberInfo = pathItem.Member,
                    Hidden = pathItem.Hidden,
                    NumberFormat = pathItem.NumberFormat,
                    TotalsRowFunction = pathItem.TotalsRowFunction,
                    TotalsRowNumberFormat = pathItem.TotalRowsNumberFormat,
                    TotalsRowLabel = pathItem.TotalRowLabel,
                    TotalsRowFormula = pathItem.TotalRowFormula,
                };
                result.Add(col);
            }
            var formulaColumnAttributes = typeof(T).FindAttributesOfType<EpplusFormulaTableColumnAttribute>();
            if (formulaColumnAttributes != null && formulaColumnAttributes.Any())
            {
                foreach (var attr in formulaColumnAttributes)
                {
                    result.Add(new ColumnInfo
                    {
                        Path = new FormulaColumnMemberPath(attr),
                        Header = attr.Header,
                        Formula = attr.Formula,
                        FormulaR1C1 = attr.FormulaR1C1,
                        NumberFormat = attr.NumberFormat,
                        TotalsRowFunction = attr.TotalsRowFunction,
                        TotalsRowNumberFormat = attr.TotalsRowNumberFormat
                    });
                }
            }
            result.ReindexAndSortColumns();
            return result;
        }
    }
}
