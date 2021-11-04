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
        public LoadFromCollectionColumns(BindingFlags bindingFlags)
        {
            _bindingFlags = bindingFlags;
        }

        private readonly BindingFlags _bindingFlags;

        internal List<ColumnInfo> Setup()
        {
            var result = new List<ColumnInfo>();
            bool sort=SetupInternal(typeof(T), result, null);
            if (sort)
            {
                ReindexAndSortColumns(result);
            }
            return result;
        }

        private bool SetupInternal(Type type, List<ColumnInfo> result, List<int> sortOrderList, string path = null)
        {
            var sort = false;
            var members = type.GetProperties(_bindingFlags);
            if (type.HasMemberWithPropertyOfType<EpplusTableColumnAttribute>())
            {
                sort = true;
                foreach (var member in members)
                {
                    if (member.HasPropertyOfType<EpplusIgnore>())
                    {
                        continue;
                    }
                    var memberPath = path != null ? $"{path}.{member.Name}" : member.Name;
                    if (member.HasPropertyOfType<EpplusNestedTableColumnAttribute>())
                    {
                        var nestedTableAttr = member.GetFirstAttributeOfType<EpplusNestedTableColumnAttribute>();
                        var attrOrder = nestedTableAttr.Order;
                        if(sortOrderList == null)
                        {
                            sortOrderList = new List<int>
                            {
                                attrOrder
                            };
                        }
                        else
                        {
                            sortOrderList.Add(attrOrder);
                        }
                        SetupInternal(member.PropertyType, result, sortOrderList, memberPath);
                        sortOrderList.RemoveAt(sortOrderList.Count - 1);
                        continue;
                    }
                    var header = default(string);
                    var sortOrder = 0;
                    var numberFormat = string.Empty;
                    var rowFunction = RowFunctions.None;
                    var totalsRowNumberFormat = string.Empty;
                    var totalsRowLabel = string.Empty;
                    var totalsRowFormula = string.Empty;
                    var colInfoSortOrderList = new List<int>();
                    var epplusColumnAttr = member.GetFirstAttributeOfType<EpplusTableColumnAttribute>();
                    if (epplusColumnAttr != null)
                    {
                        if(!string.IsNullOrEmpty(epplusColumnAttr.Header))
                        {
                            header = epplusColumnAttr.Header;
                        }
                        sortOrder = epplusColumnAttr.Order;
                        
                        if(sortOrderList != null && sortOrderList.Any())
                        {
                            colInfoSortOrderList.AddRange(sortOrderList);
                        }
                        colInfoSortOrderList.Add(epplusColumnAttr.Order);
                        numberFormat = epplusColumnAttr.NumberFormat;
                        rowFunction = epplusColumnAttr.TotalsRowFunction;
                        totalsRowNumberFormat = epplusColumnAttr.TotalsRowNumberFormat;
                        totalsRowLabel = epplusColumnAttr.TotalsRowLabel;
                        totalsRowFormula = epplusColumnAttr.TotalsRowFormula;
                    }
                    result.Add(new ColumnInfo
                    {
                        Header = string.IsNullOrEmpty(header) ? member.Name : header,
                        SortOrder = sortOrder,
                        SortOrderLevels = colInfoSortOrderList,
                        MemberInfo = member,
                        NumberFormat = numberFormat,
                        TotalsRowFunction = rowFunction,
                        TotalsRowNumberFormat = totalsRowNumberFormat,
                        TotalsRowLabel = totalsRowLabel,
                        TotalsRowFormula = totalsRowFormula,
                        Path = memberPath
                    });
                }
            }
            else
            {
                var index = 0;
                result.AddRange(members.Select(x => new ColumnInfo { Index = index++, MemberInfo = x }));
            }
            var formulaColumnAttributes = type.FindAttributesOfType<EpplusFormulaTableColumnAttribute>();
            if (formulaColumnAttributes != null && formulaColumnAttributes.Any())
            {
                sort = true;
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
            }
            return sort;
        }

        private static void ReindexAndSortColumns(List<ColumnInfo> result)
        {
            var index = 0;
            //result.Sort((a, b) => a.SortOrder.CompareTo(b.SortOrder));
            result.Sort((a, b) =>
            {
                var so1 = a.SortOrderLevels;
                var so2 = b.SortOrderLevels;
                if (so1 == null || so2 == null)
                {
                    if(a.SortOrder == b.SortOrder)
                    {
                        return a.Index.CompareTo(b.Index);
                    }
                    else
                    {
                        return a.SortOrder.CompareTo(b.SortOrder);
                    }
                }
                else if (!so1.Any() || !so2.Any())
                {
                    return a.SortOrder.CompareTo(b.SortOrder);
                }
                else
                {
                    var maxIx = so1.Count < so2.Count ? so1.Count : so2.Count;
                    for(var ix = 0; ix < maxIx; ix++)
                    {
                        var aVal = so1[ix];
                        var bVal = so2[ix];
                        if (aVal.CompareTo(bVal) == 0) continue;
                        return aVal.CompareTo(bVal);
                    }
                    return 0;
                }
            });
            result.ForEach(x => x.Index = index++);
        }        
    }
}
