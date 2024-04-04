using OfficeOpenXml.Attributes;
using OfficeOpenXml.LoadFunctions.ReflectionHelpers;
using OfficeOpenXml.Table;
using System;
/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/10/2020         EPPlus Software AB       EPPlus 5.5
 *************************************************************************************************/
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;

namespace OfficeOpenXml.LoadFunctions
{
    [DebuggerDisplay("Header: {Header}, SortOrders: {GetSortOrder()}, Index: {Index}")]
    internal class ColumnInfo
    {
        #region Constructors
        public ColumnInfo()
        {
            TotalsRowFunction = RowFunctions.None;
        }

        public ColumnInfo(EpplusFormulaTableColumnAttribute attr)
        {
            Path = new FormulaColumnMemberPath(attr);
            Header = attr.Header;
            Formula = attr.Formula;
            FormulaR1C1 = attr.FormulaR1C1;
            NumberFormat = attr.NumberFormat;
            TotalsRowFunction = attr.TotalsRowFunction;
            TotalsRowNumberFormat = attr.TotalsRowNumberFormat;
        }

        public ColumnInfo(MemberPath path)
        {
            var pathItem = path.Last();
            Header = path.GetHeader();
            Path = path;
            IsDictionaryProperty = pathItem.IsDictionaryColumn;
            MemberInfo = pathItem.Member;
            Hidden = pathItem.Hidden;
            NumberFormat = pathItem.NumberFormat;
            TotalsRowFunction = pathItem.TotalsRowFunction;
            TotalsRowNumberFormat = pathItem.TotalRowsNumberFormat;
            TotalsRowLabel = pathItem.TotalRowLabel;
            TotalsRowFormula = pathItem.TotalRowFormula;
        }

        #endregion


        public bool IsDictionaryProperty { get; set; }

        public string DictinaryKey { get; set; }
        public int Index { get; set; }

        public MemberInfo MemberInfo { get; set; }

        public string Formula { get; set; }

        public string FormulaR1C1 { get; set; }

        public string Header { get; set; }

        public bool Hidden { get; set; }

        public string NumberFormat { get; set; }

        public bool UseBuiltInHyperlinkStyle { get; set; }

        public RowFunctions TotalsRowFunction { get; set; }

        public string TotalsRowFormula { get; set; }

        public string TotalsRowNumberFormat { get; set; }

        public string TotalsRowLabel { get; set; }

        public MemberPathBase Path { get; set; }

        public override string ToString()
        {
            if(!string.IsNullOrEmpty(Header))
            {
                return Header;
            }
            return base.ToString();
        }

        public string GetSortOrder()
        {
            return Path.GetSortOrderString();
        }
    }
}
