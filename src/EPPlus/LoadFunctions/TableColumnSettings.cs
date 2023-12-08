/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  12/7/2023         EPPlus Software AB       EPPlus 7.0.4
 *************************************************************************************************/
using OfficeOpenXml.Table;
using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace OfficeOpenXml.LoadFunctions
{
    internal class TableColumnSettings
    {
        public TableColumnSettings()
        {
            
        }
        public void SetProperties(EpplusTableColumnAttribute attr)
        {
            Header = attr.Header;
            Hidden = attr.Hidden;
            NumberFormat = attr.NumberFormat;
            TotalsRowFunction = attr.TotalsRowFunction;
            TotalRowsNumberFormat = attr.TotalsRowNumberFormat;
            TotalRowLabel = attr.TotalsRowLabel;
            TotalRowFormula = attr.TotalsRowFormula;
        }

        public string GetHeader(string headerPrefix, MemberInfo member)
        {
            if (!string.IsNullOrEmpty(headerPrefix))
            {
                var header = string.IsNullOrEmpty(Header) ? member.Name : Header;
                return $"{headerPrefix} {header}";
            }
            else
            {
                return string.IsNullOrEmpty(Header) ? member.Name : Header;
            }
        }
        public string Header { get; set; }

        public bool Hidden { get; set; }

        public string NumberFormat { get; set; }

        public RowFunctions TotalsRowFunction { get; set; }

        public string TotalRowsNumberFormat { get; set; }

        public string TotalRowLabel { get; set; }

        public string TotalRowFormula { get; set; }

        public static TableColumnSettings Default
        {
            get
            {
                return new TableColumnSettings
                {
                    Header = default,
                    NumberFormat = string.Empty,
                    TotalsRowFunction = RowFunctions.None,
                    TotalRowsNumberFormat = string.Empty,
                    TotalRowLabel = string.Empty,
                    TotalRowFormula = string.Empty
                };
            }
        }

    }
}
